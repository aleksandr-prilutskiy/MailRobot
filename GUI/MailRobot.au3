#Region Header
#comments-start

	Title:			Mail Robot
	Filename:		MailRobot.au3
	Description:	Робот обработки сообщений e-mail
	Version:		0.1.0
	Date:			02.02.2022

#comments-end
#EndRegion Header

#Region Initialization
#pragma compile(ProductName, PRO100ROBOT - Mail Robot)
#pragma compile(FileVersion, 0.0.1.0)
#pragma compile(LegalCopyright, (c) 2022 PRO100ROBOT)
#pragma compile(Out, ..\bin\MailRobotGUI.exe)
#pragma compile(Icon, PRO100ROBOT.ico)
#pragma compile(x64, true)
#pragma compile(UPX, false)
#pragma compile(Console, false)
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <Encoding.au3>
Opt("GUIOnEventMode", 1)
Opt("TrayIconHide", 1)

Global	$sAppTempDir			= ''								; каталог размещения временных файлов
Global	$fCheckboxRunWeb		= True								; режим запуска веб-сервера после установки
$sAppName		= 'Робот обработки сообщений e-mail'				; название устанавливаемой программы
$sAppShortName	= 'Mail Robot'										; краткое название программы
$sAppVersion	= '0.1.0'											; версия устанавливаемой программы
$sAppPublisher	= 'PRO100ROBOT'
$sWebSiteURL	= "https://pro100robot.com"

; переменные пользовательского интерфейса
Global	$sWindowName			= 'PRO100ROBOT - Робот обработки сообщений e-mail' ; заголовок окна установщика
Global	$WindowWidth			= 500								; ширина окна программы
Global	$WindowHeight			= 143								; высота окна программы
Global	$GUI_MainWindow 											; главное окно программы
Global	$GUI_ButtonGo
Global	$sIniFileName			= @ScriptDir & "\MailRobot.ini"		; имя файла конфигурации программы

Global	$sMessage				= ""

#EndRegion Initialization
LoadConfigFile()
CreateGUI()

#Region GUI
;----------------------------------------- НАСТРОЙКА ПОЛЬЗОВАТЕЛЬСКОГО ИНТЕРФЕЙСА -------------------------------

; #FUNCTION# ====================================================================================================
; Name...........:	LoadConfigFile
; Description....:	Загрузка настроек программы из файла конфигурации.
; Syntax.........:	LoadConfigFile()
; ===============================================================================================================
Func LoadConfigFile()
 $sMessage			= StringReplace(IniRead($sIniFileName, "App", "Info", $sMessage), '/n', @CRLF)
EndFunc ;==>LoadConfigFile

; #FUNCTION# ====================================================================================================
; Name...........:	CreateGUI
; Description....:	Создание главного окна программы
; Syntax.........:	CreateGUI()
; ===============================================================================================================
Func CreateGUI()
 Local $i
 $GUI_MainWindow = GUICreate($sWindowName, $WindowWidth, $WindowHeight, -1, -1, BitOR($WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
 GUICtrlCreatePic('Logo.jpg', 1, 1, 163, 143)
 GUICtrlCreateLabel(_Encoding_UTF8ToANSI($sMessage), 200, 14, $WindowWidth - 200, 44, 0x0101)
 GUICtrlCreateLabel("© PRO100ROBOT 2021", 200, $WindowHeight - 75, $WindowWidth - 200, 20, 0x0101)
 GUICtrlSetOnEvent(-1, "OpenWeb")
 GUICtrlSetColor(-1, 0x0000ff)
 $GUI_ButtonGo = GUICtrlCreateButton('Запустить', 306, $WindowHeight - 50, 88, 44)
 GUICtrlSetOnEvent($GUI_ButtonGo, "_Go")
 GUICtrlCreateButton('Выход', 401, $WindowHeight - 50, 88, 44)
 GUICtrlSetOnEvent(-1, "ConfirmationExit")
 GUISetOnEvent($GUI_EVENT_CLOSE, 'ConfirmationExit')
 GUISetState()
 Do
  Sleep(100)
 Until False
EndFunc ;==>CreateGUI

; #FUNCTION# ====================================================================================================
; Name...........:	_Go
; Description....:	Запуск робота
; Syntax.........:	_Go()
; ===============================================================================================================
Func _Go()
 GUICtrlSetState($GUI_ButtonGo, $GUI_DISABLE)
 RunWait(@ScriptDir & "\MailRobot.exe", @ScriptDir, @SW_HIDE)
 GUICtrlSetState($GUI_ButtonGo, $GUI_ENABLE)
EndFunc ;==>_Go

; #FUNCTION# ====================================================================================================
; Name...........:	OpenWeb
; Description....:	Запуск браузера и открытие страницы программы в Интернете
; Syntax.........:	OpenWeb()
; ===============================================================================================================
Func OpenWeb()
 ShellExecute("explorer", $sWebSiteURL)
EndFunc ;==>OpenWeb

; #FUNCTION# ====================================================================================================
; Name...........:	_СonfirmationExit
; Description....:	Запрос подтверждения выхода из программы
; Syntax.........:	_СonfirmationExit()
; ===============================================================================================================
Func ConfirmationExit()
 If MsgBox(BitOR($MB_TOPMOST, $MB_ICONQUESTION, $MB_YESNO), $sWindowName, _
  'Выйти из программы?') == $IDYES Then Exit
EndFunc ;==>ConfirmationExit

#EndRegion GUI

#Requires AutoHotkey v2.0
#SingleInstance Force

;#Include %A_ScriptDir%\ahk2_lib-master\JSON.ahk
;#Include %A_ScriptDir%\ahk2_lib-master\Chrome.ahk
;#Include .\ahk2_lib-master\JSON.ahk
;#Include .\ahk2_lib-master\Chrome.ahk
;#Include .\ShotCutAutoUtil.ahk
#Include .\ShotCutAutoFunc.ahk

^r::Reload

^+m::DebugMsgToggle()  ; Ctrl + Shift + M

^+H::KeyHistory  ; Ctrl + Shift + H

^!NumpadEnter::RefreshAlcoholNormal()  ; Ctrl + Alt + NumpadEnter

^!+NumpadEnter::RefreshAlcoholNormal(true)  ; Ctrl + Alt + Shift + NumpadEnter

^NumpadEnter::RefreshAlcoholCDP(true) ; Ctrl + NumpadEnter

^+NumpadEnter::CompareAlcoholCDP(false) ; Ctrl + Shift + NumpadEnter

;^!+F2:: ProcessClose("chrome.exe")
^!+F2::Run '"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9223 --remote-allow-origins=* --user-data-dir=C:\ChromePython'
^!+F1::Run '"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir=C:\ChromeDebug'

^F12::RefreshBusRoute("5620")

^F11::RefreshBusRoute("5413")

^!F12::RefreshBusRoute("5620",true)

^!F11::RefreshBusRoute("5413",true)

^Numpad0::StartKeyHookNumpad0()

^Numpad1::
{
	if 0 {
		data := []
		for window in ComObject("Shell.Application").Windows
		{
			data.Push( [window.LocationName, window.LocationURL])
		}
		for row in data
		{
			Log( "1, Name: " row[1] ", URL: " row[2])
		}
	} else if 0 {
		data := Map()
		for window in ComObject("Shell.Application").Windows
		{
			data[window.LocationName] := window.LocationURL
		}
		for name, url in data
		{
			Log( "2, Name: " name ", URL: " url)
		}
	} else if 1 {
		data := Map()
		for window in ComObject("Shell.Application").Windows
		{
			data[window.LocationName] := [window.LocationURL, FileGetSize(window.FullName)]
		}
		for name, info in data
		{
			Log( "3, Name: " name ", URL: " info[1] ", Size: " info[2])
		}
	} else{
		windows := ""
		for window in ComObject("Shell.Application").Windows
		{
			windows .= window.LocationName " :: " window.LocationURL "`n"
		}
		Log windows
	}
	
}



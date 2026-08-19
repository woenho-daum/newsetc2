#Requires AutoHotkey v2.0
#SingleInstance Force

;#Include %A_ScriptDir%\ahk2_lib-master\JSON.ahk
;#Include %A_ScriptDir%\ahk2_lib-master\Chrome.ahk
;#Include .\ahk2_lib-master\JSON.ahk
;#Include .\ahk2_lib-master\Chrome.ahk
#Include .\ShotCutAutoUtil.ahk

^r::Reload

^+m::  ; Ctrl + Shift + M
{
    global DebugMsg

    DebugMsg := !DebugMsg   ; true <-> false 전환

    if DebugMsg
        MsgBox "디버그 모드 ON"
    else
        MsgBox "디버그 모드 OFF"
}

^+H::  ; Ctrl + Shift + H
{
    KeyHistory
}

^!NumpadEnter::  ; Ctrl + Alt + NumpadEnter
{
    ; 현재 창(엑셀) 저장
    ExcelWin := WinExist("A")

    ; 크롬 활성화
    if WinExist("ahk_exe chrome.exe")
    {
        WinActivate("ahk_exe chrome.exe")
        WinWaitActive("ahk_exe chrome.exe",,2)

        Sleep 500

        ; 최상단
        Send "{Home}"

        Sleep 200

        Send "{F5}"

        ; 새로고침이 끝날 때까지 대기
        Sleep 800

        ; 엑셀로 복귀
        WinActivate("ahk_id " ExcelWin)
    }
}

^!+NumpadEnter::  ; Ctrl + Alt + Shift + NumpadEnter
{
    ; 현재 창(엑셀) 저장
    ExcelWin := WinExist("A")

    ; Chrome가 실행 중인지 확인
    if !WinExist("ahk_exe chrome.exe")
        return

    ; Chrome 활성화
    WinActivate("ahk_exe chrome.exe")
    WinWaitActive("ahk_exe chrome.exe")

    Sleep 300

    ; 탭 검색
    Send "^+a"
    Sleep 300

    ; 검색창 내용 삭제
    Send "^a"
    Sleep 100
    Send "{Backspace}"

    ; 탭 이름 입력
    SendText "음주측정 데이터 관리 시스템"
    Sleep 500

    ; 해당 탭으로 이동
    Send "{Enter}"

    ; 탭이 바뀔 때까지 잠시 대기
    Sleep 500

    ; 최상단
    Send "{Home}"

    Sleep 500

    ; 새로고침
    Send "{F5}"

    ; 페이지 로딩 시간
    Sleep 1500

    ; 엑셀 복귀
    WinActivate("ahk_id " ExcelWin)
}

^NumpadEnter::
{
     try
    {
        ; 현재 창(엑셀) 저장
        ;ExcelWin := WinExist("A")

        objChrome := Chrome()

        ; 제목에 RouteTitle이 포함된 탭 찾기
        page := objChrome.GetPageByTitle("음주측정 데이터 관리 시스템", "contains")
        if !page
            throw Error("'음주측정 데이터 관리 시스템' 탭을 찾을 수 없습니다.CDP")

		; 1. CDP 레벨에서 탭을 해당 창의 최상단으로 (창 내부 탭 전환)
		page.Call("Page.bringToFront")
		page.Activate()
		page.WaitForLoad()

		; 2. 실제 OS 창을 찾아서 최상위로 올리기
		targetHwnd := FindChromeWindowByTitle("음주측정 데이터 관리 시스템")
		if targetHwnd {
			WinActivate("ahk_id " targetHwnd)
			WinWaitActive("ahk_id " targetHwnd,, 2)
		} else {
			throw Error("'음주측정 데이터 관리 시스템' Chrome 창을 찾을 수 없습니다.윈도우핸들")
		}

        ;---button 목록
        ;ListButtons(page)   ; 여기서 결과를 보고 정확한 id 확인

        result := CallJsFunction(page, "goresult")

        Log( "=== 음주측정갱신 ===`n" result["value"] )

        ; 엑셀 복귀
        
        SetTitleMatchMode "RegEx"

        ExcelWin := WinExist(".*-[0123][0-9]음주.*\.xlsx?.*")

		if !ExcelWin
    		throw Error("'음주측정대장' 엑셀창을 찾을 수 없습니다.")

		WinActivate("ahk_id " ExcelWin)

		page.WaitForLoad()
		
		Sleep 300

		rowCount := GetAlcoholTableRowCount(page)
		colCount := GetAlcoholTableColumnCount(page)

		Log("Row = " rowCount)
		Log("Column = " colCount)

		; Excel의 Cells(1, 7)과 비슷한 개념
		driver := GetAlcoholTableCell(page, 1, 7)
		Log("운전자 = " driver)

		xlApp := GetXlAppFromHwnd(ExcelWin)
		if !xlApp
    		throw Error("'음주측정대장' 엑셀창에서 엑셀핸들을 찾을 수 없습니다.")

		if (IsDebugging()){
			try xlApp2 := ComObjActive("Excel.Application")   ; 활성 엑셀 인스턴스 가져오기
			hwnd1 := xlApp ? xlApp.Hwnd : "없음"
			hwnd2 := xlApp2 ? xlApp2.Hwnd : "없음"

			name1 := xlApp ? xlApp.ActiveWorkbook.Name : "없음"
			name2 := xlApp2 ? xlApp2.ActiveWorkbook.Name : "없음"

			sameInstance := (xlApp && xlApp2 && hwnd1 = hwnd2) ? "동일" : "다름"

			Log("excel instance 비교, "
				. "xlApp(Hwnd=" hwnd1 ", Name=" name1 ") / "
				. "xlApp2(Hwnd=" hwnd2 ", Name=" name2 ") -> " sameInstance)
		}
		xlSheet := xlApp.ActiveSheet                 ; 현재 활성 시트

		; LookIn:=xlValues(-4163), LookAt:=xlWhole(1) → 값이 완전히 일치하는 셀 검색
		;foundCell := xlSheet.Cells.Find(driver, , -4163, 1)
		; LookIn=xlValues, LookAt=xlWhole, SearchOrder=xlByRows(1), SearchDirection=xlNext(1), MatchCase=false
		foundCell := xlSheet.Cells.Find(driver, , -4163, 1, 1, 1, false)

		if IsObject(foundCell) {
			foundCell.Select()      ; 셀 선택
			foundCell.Activate()    ; 액티브 셀로 지정 (커서 이동)
		} else {
			MsgBox( "운전자 '" driver "'를 찾을 수 없습니다.")
		}

        return true
    }
    catch Error as e
    {
        Log(
            "Message : " e.Message
            . "`nWhat : " e.What
            . "`nLine : " e.Line
            . "`nFile : " e.File
            . "`nExtra : " e.Extra
            . "`n"
        )

        return false
    }

}

^+NumpadEnter::
{
     try
    {
        ; 현재 창(엑셀) 저장
        ;ExcelWin := WinExist("A")

        objChrome := Chrome()

        ; 제목에 RouteTitle이 포함된 탭 찾기
        page := objChrome.GetPageByTitle("음주측정 데이터 관리 시스템", "contains")
        if !page
            throw Error("'음주측정 데이터 관리 시스템' 탭을 찾을 수 없습니다.")

		; 1. CDP 레벨에서 탭을 해당 창의 최상단으로 (창 내부 탭 전환)
		page.Call("Page.bringToFront")
		page.Activate()
		page.WaitForLoad()

		; 2. 실제 OS 창을 찾아서 최상위로 올리기
		targetHwnd := FindChromeWindowByTitle("음주측정 데이터 관리 시스템")
		if targetHwnd {
			WinActivate("ahk_id " targetHwnd)
			WinWaitActive("ahk_id " targetHwnd,, 2)
		} else {
			throw Error("'음주측정 데이터 관리 시스템' Chrome 창을 찾을 수 없습니다.")
		}
        ;---button 목록
        ;ListButtons(page)   ; 여기서 결과를 보고 정확한 id 확인

        ;result := CallJsFunction(page, "goresult")
        ;Log( "=== 음주측정갱신 ===`n" result["value"] )

        ; 엑셀 복귀
        
        SetTitleMatchMode "RegEx"

        ExcelWin := WinExist(".*-[0123][0-9]음주.*\.xlsx?.*")

        if !ExcelWin
    		throw Error("'음주측정대장' 엑셀창을 찾을 수 없습니다.")
        
		WinActivate("ahk_id " ExcelWin)

		page.WaitForLoad()
		
		Sleep 300

		rowCount := GetAlcoholTableRowCount(page)
		colCount := GetAlcoholTableColumnCount(page)

		Log("Row = " rowCount ", rowCount 타입 = " Type(rowCount))
		Log("Column = " colCount)

		;xlApp := ComObjActive("Excel.Application")   ; 활성 엑셀 인스턴스 가져오기
		xlApp := GetXlAppFromHwnd(ExcelWin)
		if !xlApp
			throw Error("'음주측정대장' 엑셀창에서 엑셀핸들을 찾을 수 없습니다.")

		; 엑셀 업데이트 잠시 보류처리시 오류발생하면 복구 해야한다
		try{
			xlSheet := xlApp.ActiveSheet                 ; 현재 활성 시트

			; 만약 AHK가 마우스/키보드로 직접 엑셀 창을 조작(클릭, 타이핑 등)하는 방식이라면 이 방법은 소용없다.. ㅠㅠ
			xlApp.Interactive := false      ; 사용자 입력(마우스/키보드) 비활성화
			xlApp.ScreenUpdating := false   ; 화면 업데이트 끄기
			xlApp.EnableEvents := false     ; (선택) 매크로/이벤트 발생 억제
			xlApp.Calculation := -4135      ; (선택) xlCalculationManual, 수식 자동계산도 끄기

			nAlcoholCnt := 0
			Loop rowCount
			{
				row := A_Index
				driver := GetAlcoholTableCell(page, row, 7)
				
				; LookIn:=xlValues(-4163), LookAt:=xlWhole(1) → 값이 완전히 일치하는 셀 검색
				;foundCell := xlSheet.Cells.Find(driver, , -4163, 1)
				; LookIn=xlValues, LookAt=xlWhole, SearchOrder=xlByRows(1), SearchDirection=xlNext(1), MatchCase=false
				foundCell := xlSheet.Cells.Find(driver, , -4163, 1, 1, 1, false)
			
				if IsObject(foundCell) {
					; G열인지 확인
					if foundCell.Column = 7
					{
						; G열에서 왼쪽으로 3칸 → D열
						value := foundCell.Offset(0, -3).Value

						Log("운전자 = " driver ", 찾은 셀 = " foundCell.Address ", 왼쪽 3번째 값 = " value)

						if value = "본사"
						{
							nAlcoholCnt++
							foundCell.Interior.Color := 0xFFFFCC
							foundCell.Offset(0, -1).Interior.Color := 0xFFFFCC
							foundCell.Offset(0, 1).Interior.Color := 0XFFCC99 ;0xFFFFCC 
						}
					}
					else
					{
						Log("운전자 '" driver "'는 G열이 아닙니다.")
						Log("찾은 셀 = " foundCell.Address ",column = " foundCell.column)
					}
				} else {
					Log( "운전자 '" driver "'를 찾을 수 없습니다.")
				}
			}

		} catch as e {
			Log("에러 발생: " e.Message)
		} finally {
			xlApp.Calculation := -4105      ; (껐다면)xlCalculationAutomatic
			xlApp.EnableEvents := true      ; (껐다면)매크로/이벤트 발생 활성화
			xlApp.ScreenUpdating := true    ; 화면 업데이트 켜기
			xlApp.Interactive := true       ; 사용자 입력(마우스/키보드) 다시 활성화
		}

		WinActivate("ahk_id " ExcelWin)
		WinWaitActive("ahk_id " ExcelWin)

		global ex, ey, ew, eh, msgTitle
		; 엑셀 창의 위치/크기 가져오기
		WinGetPos(&ex, &ey, &ew, &eh, "ahk_id " ExcelWin)

		if(1)
		{
			msgTitle := "음주측정 확인"  ; MsgBox 제목 (구분용)

			; MsgBox가 뜨는 걸 감지해서 중앙으로 이동시키는 타이머 시작
			SetTimer(CenterMsgBox, 20)

			MsgBox("음주측정 미확인자 " nAlcoholCnt "명", msgTitle)

		}else if(1){
			ShowCenteredMsg(ExcelWin, "음주측정 미확인자 " nAlcoholCnt "명")
		}else if(0){
			text := "음주측정 미확인자 " nAlcoholCnt "명"

			DllCall("MessageBox",
				"Ptr", ExcelWin,
				"Str", text,
				"Str", "알림",
				"UInt", 0x40)   ; MB_ICONINFORMATION
		}else if(0){
			MsgBox ("음주측정 미확인자 " nAlcoholCnt "명")
		}else{
			myGui := Gui("+Owner" ExcelWin, "알림")
			myGui.AddText(, "음주측정 미확인자 " nAlcoholCnt "명")
			myGui.AddButton("Default", "확인").OnEvent("Click", (*) => myGui.Destroy())
			myGui.Show()
		}
        
        return true
    }
    catch Error as e
    {
        Log(
            "Message : " e.Message
            . "`nWhat : " e.What
            . "`nLine : " e.Line
            . "`nFile : " e.File
            . "`nExtra : " e.Extra
            . "`n"
        )

        return false
    }

}

; Ctrl + Alt + Shift + F2
; 모든 Chrome 종료
^!+F2::
{
    ProcessClose("chrome.exe")
    
}


; Ctrl + Alt + Shift + F1
; 크롬디버그포트 실행
^!+F1::
{
    ; 혹시 남아있는 Chrome 종료
    ;ProcessClose("chrome.exe")

    ; 종료될 때까지 잠시 대기
    Sleep 1000

    Run '"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir=C:\ChromeDebug'
}

^F12::
{
    RefreshBusRoute("5620")
}

^F11::
{
    RefreshBusRoute("5413")
}

^!F12::
{
    RefreshBusRoute("5620",true)
}

^!F11::
{
    RefreshBusRoute("5413",true)
}

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

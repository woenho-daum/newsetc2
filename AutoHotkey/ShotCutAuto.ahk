#Requires AutoHotkey v2.0
#SingleInstance Force

;#Include %A_ScriptDir%\ahk2_lib-master\JSON.ahk
;#Include %A_ScriptDir%\ahk2_lib-master\Chrome.ahk
#Include .\ahk2_lib-master\JSON.ahk
#Include .\ahk2_lib-master\Chrome.ahk

^!F12::
{
    KeyHistory
}

^!NumpadEnter::
{
    ; 현재 창(엑셀) 저장
    ExcelWin := WinExist("A")

    ; 크롬 활성화
    if WinExist("ahk_exe chrome.exe")
    {
        WinActivate("ahk_exe chrome.exe")
        WinWaitActive("ahk_exe chrome.exe",,2)

        Sleep 200
        Send "{F5}"

        ; 새로고침이 끝날 때까지 대기
        Sleep 800

        ; 엑셀로 복귀
        WinActivate("ahk_id " ExcelWin)
    }
}

^!+NumpadEnter::
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
    Sleep 700

    ; 새로고침
    Send "{F5}"

    ; 페이지 로딩 시간
    Sleep 1500

    ; 엑셀 복귀
    WinActivate("ahk_id " ExcelWin)
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
    ProcessClose("chrome.exe")

    ; 종료될 때까지 잠시 대기
    Sleep 1000

    Run '"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir=C:\ChromeDebug'
}

RefreshBusRoute(RouteTitle)
{
    try
    {
        chrome := Chrome()

        ; 제목에 RouteTitle이 포함된 탭 찾기
        page := chrome.GetPageByTitle(RouteTitle, "contains")
        if !page
            throw Error("'" RouteTitle "' 탭을 찾을 수 없습니다.")

        ; 탭 활성화(선택사항)
        page.Activate()
        ; 로딩 완료 대기
        page.WaitForLoad()

        ; 새로고침 버튼 클릭
        page.Evaluate("
        (
            document.getElementById('busRouteRefresh').click();
        )")
        return true
    }
    catch Error as e
    {
        MsgBox
        (
            "Message : " e.Message
            . "`n`nWhat : " e.What
            . "`n`nLine : " e.Line
            . "`n`nFile : " e.File
            . "`n`nExtra : " e.Extra
        )
        return false
    }
}

^!+F12::
{

    ExcelWin := WinExist("A")

    RefreshBusRoute("5620")

    if ExcelWin
        WinActivate("ahk_id " ExcelWin)
}

; 부팅(또는 스크립트 실행) 후 5초 뒤 핫스팟 ON
;SetTimer AutoHotspot, -5000

; Ctrl + Alt + NumPad+ 로도 실행 가능
;^!NumpadAdd::AutoHotspot()

;AutoHotspot()
;{
;	;MsgBox "단축키가 눌렸습니다."
;    Run 'powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "' A_ScriptDir '\HotspotOn.ps1"',, "Hide"
;	;Run 'powershell.exe -WindowStyle Normal -NoExit -ExecutionPolicy Bypass -File "' A_ScriptDir '\HotspotOn.ps1"'
;}


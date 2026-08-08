#Requires AutoHotkey v2.0

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

#Requires AutoHotkey v2.0

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

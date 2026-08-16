#Requires AutoHotkey v2.0
#SingleInstance Force

;#Include %A_ScriptDir%\ahk2_lib-master\JSON.ahk
;#Include %A_ScriptDir%\ahk2_lib-master\Chrome.ahk
#Include .\ahk2_lib-master\JSON.ahk
#Include .\ahk2_lib-master\Chrome.ahk

global DebugMsg := false

^+m::  ; Ctrl + Alt + M
{
    global DebugMsg

    DebugMsg := !DebugMsg   ; true <-> false 전환

    if DebugMsg
        Log "디버그 모드 ON"
    else
        Log "디버그 모드 OFF"
}

^+H::
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

CallJsFunction(page,funcName) {
    js :=
    (
        Format("
        (
            function(){{
                if (typeof {1} !== 'function')
                    return '{1}:function-not-found';

                {1}();
                return '{1}:called';
            }})();
        ", funcName)
    )

    result := page.Evaluate(js)
    return result
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
            throw Error("'음주측정 데이터 관리 시스템' 탭을 찾을 수 없습니다.")

        ; 탭 활성화(선택사항)
        page.Activate()
        ; 로딩 완료 대기
        page.WaitForLoad()

        ;---button 목록
        ;ListButtons(page)   ; 여기서 결과를 보고 정확한 id 확인

        result := CallJsFunction(page, "goresult")

        Log( "=== 음주측정갱신 ===`n" result["value"] )

        ; 엑셀 복귀
        
        SetTitleMatchMode "RegEx"

        ExcelWin := WinExist(".*-[0123][0-9]음주\.xlsx?.*")

        if ExcelWin
            WinActivate("ahk_id " ExcelWin)

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
    ProcessClose("chrome.exe")

    ; 종료될 때까지 잠시 대기
    Sleep 1000

    Run '"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir=C:\ChromeDebug'
}


;----- 시작하면 디버그코드인지 확인하자---
DllCall("GetCommandLine", "str")

IsDebugging() {
    static result := ""
    if (result = "")
        result := InStr(DllCall("GetCommandLine", "str"), "/Debug=") ? true : false
    return result
}

Log(msg) {
    
    global DebugMsg

    if DebugMsg{
        if IsDebugging()
            OutputDebug(msg "`n")
        else
            MsgBox(msg)
    }
}

; ---- 페이지(및 iframe)에 있는 버튼류 요소 id/value 나열 ----
ListButtons(page)
{
    result := page.Evaluate("
    (
        (function(){
            function collect(doc, label){
                var out = [];
                var nodes = doc.querySelectorAll('input[type=button], button, [id]');
                nodes.forEach(function(el){
                    if (el.id) out.push(label + ' id=' + el.id + ' value=' + (el.value || el.textContent || '').substring(0,20));
                });
                return out;
            }
            var lines = collect(document, 'top');
            var frames = document.querySelectorAll('iframe');
            frames.forEach(function(f, i){
                try {
                    lines = lines.concat(collect(f.contentDocument, 'iframe' + i));
                } catch(e) { lines.push('iframe' + i + ' 접근불가: ' + e.message); }
            });
            return lines.join('\n');
        })();
    )")
    Log("=== 버튼 목록 ===`n" result["value"])
}

; ---- 요소가 활성화될 때까지 기다렸다가 클릭하는 공통 함수 ----
; page      : Chrome.ahk의 page 객체
; elementId : 클릭할 요소의 id
; timeoutLoop : 최대 대기 반복 횟수 (기본 20회, 300ms 간격 = 최대 6초)
; ---- 요소가 존재할 때까지 기다리는 함수 (존재 여부만 체크) ----
; ---- 존재확인 + 클릭을 "한 번의 Evaluate 호출"로 처리 (핵심) ----
CheckAndClickElement(page, elementId)
{
    jsTemplate := "
    (
        (function(){
            var el = document.getElementById('__ELEMENT_ID__');
            if (!el) return 'not-exist';
            if (el.classList.contains('jqx-fill-state-disabled')) return 'disabled';
            ['mousedown','mouseup','click'].forEach(function(type){
                el.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, view:window}));
            });
            return 'dispatched';
        })();
    )"

    js := StrReplace(jsTemplate, "__ELEMENT_ID__", elementId)
    result := page.Evaluate(js)
    return result["value"]
}

; ---- 재시도까지 포함한 최종 함수 ----
ClickElementWhenReady(page, elementId, maxWaitMs := 8000, intervalMs := 200)
{
    elapsed := 0
    Loop {
        status := CheckAndClickElement(page, elementId)
        if IsDebugging()
            Log("[" elementId "] elapsed=" elapsed " : " status)

        if (status = "dispatched")
            return "dispatched"

        elapsed += intervalMs
        if (elapsed >= maxWaitMs)
            return "failed"

        Sleep intervalMs
    }
}

RefreshBusRoute(RouteTitle, bGridView:=false)
{
    try
    {
        objChrome := Chrome()

        ; 제목에 RouteTitle이 포함된 탭 찾기
        page := objChrome.GetPageByTitle(RouteTitle, "contains")
        if !page
            throw Error("'" RouteTitle "' 탭을 찾을 수 없습니다.")

        ; 탭 활성화(선택사항)
        page.Activate()
        ; 로딩 완료 대기
        page.WaitForLoad()

        ; listup button
        ;ListButtons(page)   ; 여기서 결과를 보고 정확한 id 확인

        ; 새로고침 버튼 클릭
        ;result := page.Evaluate("
        ;(
        ;    (function(){
        ;        var el = document.getElementById('busRouteRefresh');
        ;        return 'el=' + (el ? 'FOUND' : 'NULL') + ', bodyHTML길이=' + document.body.innerHTML.length;
        ;    })();
        ;)")
        ;Log("직접확인: " result["value"])

        result := ClickElementWhenReady(page, "busRouteRefresh")
        Log("busRouteRefresh 최종결과: " result "`n")

        ;openBusChartGrid - 표로보기
        if bGridView {
            result := ClickElementWhenReady(page, "openBusChartGrid")
            Log("openBusChartGrid 최종결과: " result "`n")

            ;clickResult := page.Evaluate("
            ;(
            ;    (function(){
            ;        var el = document.getElementById('openBusChartGrid');
            ;        if (!el) return 'no-element';
            ;        ['mousedown','mouseup','click'].forEach(function(type){
            ;            el.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, view:window}));
            ;        });
            ;        return 'dispatched';
            ;    })();
            ;)")
            ;Log("openBusChartGrid 최종결과: " clickResult["value"] "`n")
        }

        SetTitleMatchMode "RegEx"

        ExcelWin := WinExist(".*-[0123][0-9]_배차시간표\.xlsx?.*")

        if ExcelWin
        {
            WinActivate("ahk_id " ExcelWin)
            WinWaitActive("ahk_id " ExcelWin)

            ; 실행 중인 Excel에 연결
            xl := ComObjActive("Excel.Application")

            ; 활성 통합문서의 첫 번째 시트 선택
            xl.ActiveWorkbook.Worksheets(1).Activate()
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


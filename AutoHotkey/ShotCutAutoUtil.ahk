#Requires AutoHotkey v2.0
#SingleInstance Force

;#Include %A_ScriptDir%\ahk2_lib-master\JSON.ahk
;#Include %A_ScriptDir%\ahk2_lib-master\Chrome.ahk
#Include .\ahk2_lib-master\JSON.ahk
#Include .\ahk2_lib-master\Chrome.ahk

global DebugMsg := false
global ex:=0, ey:=0, ew:=0, eh:=0, msgTitle:="알림"

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

    if IsDebugging(){
        OutputDebug(msg "`n")
    }else{
        if DebugMsg
            MsgBox(msg)
    }
}

GetAlcoholTableCell(page, row, col) {
    js := Format("
    (
        (function() {{
            var table = document.getElementById('memberList');

            if (!table)
                return 'TABLE_NOT_FOUND';

            var rows = table.querySelectorAll('tbody tr');

            if ({1} < 1 || {1} > rows.length)
                return 'ROW_OUT_OF_RANGE';

            var cells = rows[{1} - 1].querySelectorAll('td');

            if ({2} < 1 || {2} > cells.length)
                return 'COL_OUT_OF_RANGE';

            return cells[{2} - 1].innerText.trim();
        }})()
    )", row, col)

    result := page.Evaluate(js)

    return result["value"]
}


GetAlcoholTableRowCount(page) {
    js := "document.querySelectorAll('#memberList tbody tr').length"
    return page.Evaluate(js)["value"]
}


GetAlcoholTableColumnCount(page) {
    js := "
    (
        (() => {
            const row = document.querySelector('#memberList tbody tr');
            return row ? row.cells.length : 0;
        })();
    )"

    return page.Evaluate(js)["value"]
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

GetXlAppFromHwnd(hwnd) {
    if !hwnd
        return ""

    ; Excel 메인 창 -> XLDESK -> EXCEL7 순으로 자식 윈도우 탐색
    hwndChild := DllCall("FindWindowEx", "ptr", hwnd, "ptr", 0, "str", "XLDESK", "ptr", 0, "ptr")
    hwndChild := DllCall("FindWindowEx", "ptr", hwndChild, "ptr", 0, "str", "EXCEL7", "ptr", 0, "ptr")
    if !hwndChild
        return ""

    ; IID_IDispatch GUID 준비
    IID_IDispatch := Buffer(16, 0)
    DllCall("ole32\CLSIDFromString", "wstr", "{00020400-0000-0000-C000-000000000046}", "ptr", IID_IDispatch)

    ; OBJID_NATIVEOM(-16)으로 Excel의 네이티브 오브젝트 모델 획득
    result := DllCall("oleacc\AccessibleObjectFromWindow"
        , "ptr", hwndChild
        , "uint", 0xFFFFFFF0
        , "ptr", IID_IDispatch
        , "ptr*", &pdisp := 0)

    if (result != 0) || !pdisp
        return ""

    xlWindow := ComValue(9, pdisp)
    return xlWindow.Application
}

FindChromeWindowByTitle(searchText) {
    hwndList := WinGetList("ahk_exe chrome.exe")
    for hwnd in hwndList {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            if InStr(title, searchText)
                return hwnd
        }
    }
    return 0
}

CenterMsgBox() {
    global ex, ey, ew, eh, msgTitle
    msgId := WinExist(msgTitle)
    if msgId {
        WinGetPos(&mx, &my, &mw, &mh, "ahk_id " msgId)
        newX := ex + (ew - mw) / 2
        newY := ey + (eh - mh) / 2
        WinMove(newX, newY, , , "ahk_id " msgId)
        SetTimer(, 0)  ; 이동 완료 후 타이머 종료
    }
}

ShowCenteredMsg(targetWin, text, title := "알림") {
    WinGetPos(&tx, &ty, &tw, &th, "ahk_id " targetWin)

    myGui := Gui("+AlwaysOnTop", title)
    myGui.SetFont("s10")
    myGui.Add("Text", "w250", text)
    myGui.Add("Button", "w80 Default", "확인").OnEvent("Click", (*) => myGui.Destroy())

    ; 일단 화면 밖(임시위치)에 표시해서 실제 크기를 얻어냄
    myGui.Show("Hide")
    myGui.GetPos(, , &gw, &gh)

    ; 대상 창(엑셀) 중앙 좌표 계산
    newX := tx + (tw - gw) / 2
    newY := ty + (th - gh) / 2

    myGui.Show("x" newX " y" newY)
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
; timeoutLoop : 최대 대기 반복 횟수 (기본 20회, 200ms 간격 = 최대 4초)
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
ClickElementWhenReady(page, elementId, maxWaitMs := 4000, intervalMs := 200)
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

        ; 현재 활성화된 Chrome 창의 HWND
        page.Call("Page.bringToFront")
        WinActivate "ahk_exe chrome.exe"
        ; 탭 활성화(선택사항)
        page.Activate()
        ; 로딩 완료 대기
        page.WaitForLoad()

        ; listup button
        ;ListButtons(page)   ; 여기서 결과를 보고 정확한 id 확인

        ; 새로고침 버튼 클릭을 직접해보기
        ;result := page.Evaluate("
        ;(
        ;    (function(){
        ;        var el = document.getElementById('busRouteRefresh');
        ;        return 'el=' + (el ? 'FOUND' : 'NULL') + ', bodyHTML길이=' + document.body.innerHTML.length;
        ;    })();
        ;)")
        ;Log("직접확인: " result["value"])

		; 새로고침 버튼 클릭을 공통함수 사용하기
        result := ClickElementWhenReady(page, "busRouteRefresh")
        Log("busRouteRefresh 최종결과: " result "`n")

        ;openBusChartGrid - 표로보기
        if bGridView {
            result := ClickElementWhenReady(page, "openBusChartGrid")
            Log("openBusChartGrid 최종결과: " result "`n")

			; 표보기 버튼 클릭을 직접해보기
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
            ;xl := ComObjActive("Excel.Application")
			xlApp := GetXlAppFromHwnd(ExcelWin)

            if IsObject(xlApp) {
                xlApp.ActiveWorkbook.Windows(1).Activate()
                ;xl.ActiveWorkbook.Worksheets(1).Activate()
                xlApp.ActiveWorkbook.Worksheets("Tablib Dataset").Activate()
            } else {
				throw Error("해당 창의 Excel 객체를 가져오지 못했습니다.")
            }
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



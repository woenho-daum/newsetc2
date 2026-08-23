#Requires AutoHotkey v2.0
#SingleInstance Force

;#Include %A_ScriptDir%\ahk2_lib-master\JSON.ahk
;#Include %A_ScriptDir%\ahk2_lib-master\Chrome.ahk
#Include .\ahk2_lib-master\JSON.ahk
#Include .\ahk2_lib-master\Chrome.ahk

global g_DebugMsg := false
global g_ex:=0, g_ey:=0, g_ew:=0, g_eh:=0, g_msgTitle:="알림"
global g_ih := ""

;----- 시작하면 디버그코드인지 확인하자---
DllCall("GetCommandLine", "str")

IsDebugging() {
    static result := ""
    if (result = "")
        result := InStr(DllCall("GetCommandLine", "str"), "/Debug=") ? true : false
    return result
}

Log(msg) {
    
    global g_DebugMsg

	if g_DebugMsg
	{
		MsgBox(msg)
    } else {
	    OutputDebug(msg "`n")
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
    global g_ex, g_ey, g_ew, g_eh, g_msgTitle
    msgId := WinExist(msgTitle)
    if msgId {
        WinGetPos(&mx, &my, &mw, &mh, "ahk_id " msgId)
        newX := g_ex + (g_ew - mw) / 2
        newY := g_ey + (g_eh - mh) / 2
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

; ---- 페이지(및 iframe)에 있는 table 목록 ----
ListTables(page)
{
    result := page.Evaluate("
    (
        (function(){
            function collect(doc, label){
                var out = [];
                var tables = doc.querySelectorAll('table');

                tables.forEach(function(el, i){
                    if (el.id)
                        out.push(label + ' id=' + el.id);
                    else
                        out.push(label + ' table[' + i + ']');
                });

                return out;
            }

            var lines = collect(document, 'top');

            var frames = document.querySelectorAll('iframe');

            frames.forEach(function(f, i){
                try {
                    lines = lines.concat(collect(f.contentDocument, 'iframe' + i));)
                } catch(e) {
                    lines.push('iframe' + i + ' 접근불가: ' + e.message);
                }
            });

            return lines.join('\n');
        })();
    )")

    Log("=== 테이블 목록 ===`n" result["value"])
}

; ListTables(page) 결과를 가지고
; GetTable(page, 0)
; GetTable(page, 1)
; GetTable(page, 2)
; 형식으로 호출한다
; ---- 페이지(및 iframe)의 특정 table 내용 가져오기 ----
GetTable(page, index)
{
    result := page.Evaluate("
    (
        (function(targetIndex){
            var currentIndex = 0;

            function findTable(doc){
                var tables = doc.querySelectorAll('table');

                for (var i = 0; i < tables.length; i++){
                    if (currentIndex === targetIndex)
                        return tables[i];

                    currentIndex++;
                }

                return null;
            }

            var table = findTable(document);

            if (!table){
                var frames = document.querySelectorAll('iframe');

                for (var i = 0; i < frames.length; i++){
                    try {
                        if (frames[i].contentDocument){
                            table = findTable( frames[i].contentDocument);

                            if (table)
                                break;
                        }
                    } catch(e) {
                        // 접근할 수 없는 iframe은 건너뜀
                    }
                }
            }

            if (!table)
                return 'table[' + targetIndex + '] 없음';

            var out = [];

            table.querySelectorAll('tr').forEach(function(tr){
                var row = [];

                tr.querySelectorAll('th, td').forEach(function(cell){
                    row.push(
                        (cell.innerText || cell.textContent || '')
                        .replace(/\s+/g, ' ')
                        .trim() );
                });

                if (row.length > 0)
                    out.push(row.join(' | '));
            });

            return out.join('\n');

        })(" . index . ")
    )")

    Log("=== table[" . index . "] ===`n" result["value"])

    return result["value"]
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


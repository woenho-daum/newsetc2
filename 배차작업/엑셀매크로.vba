If False Then
    ' 여기에 비활성화할 코드 블록
    MsgBox "이 코드는 실행되지 않습니다."
    
    =OR(ROW()=20,
       AND(ROW()>=35,ROW()<=52),
            COLUMN()=COLUMN(H:H),
            COLUMN()=COLUMN(K:K),
            COLUMN()=COLUMN(N:N)
    )
    
	=IF(AND(G2<>"", I2<>""), (I2-G2)*1440, "")
	=IF(G2<>"",(I2-G2)*1440,"")
	
    =IF(J2<>"",(L2-J2)*1440,"")


End If

' WorkSheet 모듈
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 모든 행의 색 초기화
    Cells.Interior.ColorIndex = xlNone
    ' 현재 커서가 있는 행 색 반전
    Target.EntireRow.Interior.Color = RGB(0, 255, 255)
    Target.EntireColumn.Interior.Color = RGB(0, 255, 255)
End Sub

' RGB(166, 139, 65)
' RGB(0, 255, 255)
' RGB(194, 194, 194)

Sub ReplaceSpecificText1()

    Dim cell As Range

    For Each cell In ActiveSheet.UsedRange
        If VarType(cell.Value) = vbString Then
            cell.Value = Replace(cell.Value, "새벽" & vbLf & "주간", "새벽연장")
        End If
    Next cell

End Sub

Sub ReplaceSpecificText2()

    Dim cell As Range

    For Each cell In ActiveSheet.UsedRange
        If VarType(cell.Value) = vbString Then
            cell.Value = Replace(cell.Value, "주간" & vbLf & "야간", "야간연장")
        End If
    Next cell

End Sub

Sub 줄바꿈셀만변경()

    Dim c As Range

    For Each c In ActiveSheet.UsedRange
        If InStr(c.Value, Chr(10)) > 0 Then
            c.Value = Replace(c.Value, Chr(10), "")
        End If
    Next c

End Sub

Sub BaeChaRDB()

    Dim Wb As Workbook
    Dim ws As Worksheet
    Dim newWs As Worksheet

    Dim lastCol As Long
    Dim lastRow As Long
    Dim nCol As Long
    Dim nRow As Long
    Dim newRow As Long
    Dim nErrorCnt As Long

    On Error Resume Next
    
    Set Wb = ActiveWorkbook
        
    If Wb Is Nothing Then Exit Sub
    
    ' 시트 존재 체크
    Const sheetName As String = BaeChaBackup_NAME
    
    If SheetExists(Wb, sheetName) = False Then Exit Sub
    
    Set ws = ActiveWorkbook.Worksheets(BaeChaBackup_NAME)
    
    On Error GoTo 0
    
    If ws Is Nothing Then
        
        'MsgBox "시트를 찾을 수 없습니다. 시트 이름을 확인하세요."
        Exit Sub
    End If
    
    '=================================================
    ' 1단계 : 결과 시트 생성
    '=================================================
    
    '경고없이 존재하면 지운다
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Worksheets(BaeChaRDB_NAME).Delete
    On Error GoTo 0
    
    '경고를 하게 한다
    Application.DisplayAlerts = True
    
    Set newWs = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    newWs.Name = BaeChaRDB_NAME
    
    '헤더 작성
    newWs.Range("A1:J1") = Array("노선", "편성", "운행요일", "구분", "순번", "차량번호", "기사", "시간", "분", "차이")
    
    '=================================================
    ' 2단계 : 시간 컬럼을 세로로 변환
    '=================================================
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    newRow = 2
    
    For nRow = 2 To lastRow
    
        '시간01부터 마지막 시간컬럼까지
        For nCol = 9 To lastCol
    
            If Trim(ws.Cells(nRow, nCol).Value) <> "" Then
    
                newWs.Cells(newRow, 1) = ws.Cells(nRow, 2) '노선
                newWs.Cells(newRow, 2) = ws.Cells(nRow, 3) '편성
                newWs.Cells(newRow, 3) = ws.Cells(nRow, 4) '운행요일
                newWs.Cells(newRow, 4) = ws.Cells(nRow, 5) '구분
                newWs.Cells(newRow, 5) = ws.Cells(nRow, 6) '순번
                newWs.Cells(newRow, 6) = ws.Cells(nRow, 7) '차량번호
                newWs.Cells(newRow, 7) = ws.Cells(nRow, 8) '기사
                newWs.Cells(newRow, 8) = ws.Cells(nRow, nCol) '시간
                newWs.Cells(newRow, 9).Value = TimeValue(ws.Cells(nRow, nCol).Value) * 1440 '분
    
                newRow = newRow + 1
    
            End If
        Next nCol
    
    Next nRow

    '=================================================
    ' 3단계 : 시간순으로 소트
    '=================================================
    
    lastRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).row
    lastCol = newWs.Cells(1, newWs.Columns.Count).End(xlToLeft).Column
    
    With newWs.Sort
        .SortFields.Clear
    
        .SortFields.Add key:=newWs.Range("A2:A" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
    
        .SortFields.Add key:=newWs.Range("D2:D" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
    
        .SortFields.Add key:=newWs.Range("H2:H" & lastRow), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
    
        .SetRange newWs.Range(newWs.Cells(1, 1), newWs.Cells(lastRow, lastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '=================================================
    ' 4단계 : 차량그룹을 최소 순번순으로 재배치
    '=================================================
    '=================================================
    ' 4단계 : 차량그룹을 최소 순번순으로 재배치
    '=================================================

    Dim arrData As Variant
    Dim outData As Variant
    Dim nData As Long
    Dim i As Long, j As Long, k As Long, a As Long, b As Long
    Dim blockCount As Long
    Dim grpRouteEnd As Long
    Dim tmpVeh As Variant
    Dim minSeq As Double
    Dim ts As Long, te As Long, tm As Double, tr As Variant

    ' 블록 정보 저장용 배열 (시작행, 끝행, 최소순번, 노선)
    Dim blkStart() As Long
    Dim blkEnd() As Long
    Dim blkMin() As Double
    Dim blkRoute() As Variant

    nData = lastRow - 1  ' 데이터 행수 (헤더 제외)

    If nData > 0 Then

        arrData = newWs.Range(newWs.Cells(2, 1), newWs.Cells(lastRow, lastCol)).Value

        ReDim outData(1 To nData, 1 To lastCol)
        ReDim blkStart(1 To nData)
        ReDim blkEnd(1 To nData)
        ReDim blkMin(1 To nData)
        ReDim blkRoute(1 To nData)

        ' ---- 4-1. 블록(같은 노선 + 같은 차량번호) 구분 ----
        blockCount = 0
        i = 1

        Do While i <= nData

            blockCount = blockCount + 1
            blkStart(blockCount) = i
            blkRoute(blockCount) = arrData(i, 1)   '노선
            tmpVeh = arrData(i, 6)                 '차량번호
            minSeq = val(arrData(i, 5))             '순번(숫자변환) 초기값

            j = i

            Do While j + 1 <= nData
                If arrData(j + 1, 1) = blkRoute(blockCount) And arrData(j + 1, 6) = tmpVeh Then
                    j = j + 1
                    If val(arrData(j, 5)) < minSeq Then minSeq = val(arrData(j, 5))
                Else
                    Exit Do
                End If
            Loop

            blkEnd(blockCount) = j
            blkMin(blockCount) = minSeq

            i = j + 1

        Loop

        ' ---- 4-2. 같은 노선 구간별로 블록을 최소순번 기준 오름차순 안정정렬 ----
        k = 1

        Do While k <= blockCount

            ' 같은 노선인 블록 범위 찾기
            grpRouteEnd = k
            Do While grpRouteEnd + 1 <= blockCount
                If blkRoute(grpRouteEnd + 1) = blkRoute(k) Then
                    grpRouteEnd = grpRouteEnd + 1
                Else
                    Exit Do
                End If
            Loop

            ' 삽입정렬 (안정정렬 - 동일 최소순번이면 원래 순서 유지)
            For a = k + 1 To grpRouteEnd
                ts = blkStart(a): te = blkEnd(a): tm = blkMin(a): tr = blkRoute(a)
                b = a - 1
                Do While b >= k
                    If blkMin(b) > tm Then
                        blkStart(b + 1) = blkStart(b)
                        blkEnd(b + 1) = blkEnd(b)
                        blkMin(b + 1) = blkMin(b)
                        blkRoute(b + 1) = blkRoute(b)
                        b = b - 1
                    Else
                        Exit Do
                    End If
                Loop
                blkStart(b + 1) = ts
                blkEnd(b + 1) = te
                blkMin(b + 1) = tm
                blkRoute(b + 1) = tr
            Next a

            k = grpRouteEnd + 1

        Loop

        ' ---- 4-3. 정렬된 블록 순서대로 outData 재구성 ----
        newRow = 0

        For k = 1 To blockCount
            For i = blkStart(k) To blkEnd(k)
                newRow = newRow + 1
                For nCol = 1 To lastCol
                    outData(newRow, nCol) = arrData(i, nCol)
                Next nCol
            Next i
        Next k

        ' ---- 4-4. 결과를 시트에 재기록 ----
        newWs.Range(newWs.Cells(2, 1), newWs.Cells(lastRow, lastCol)).Value = outData

    End If
    

    
    
    '=================================================
    ' 5단계 : 시간 차이값 계산 하여 그 차이가 일정 수준 이하면 오류 처리
    '=================================================
    
    nErrorCnt = 0
    
    For newRow = 2 To lastRow
        If newWs.Cells(newRow - 1, 1).Value = newWs.Cells(newRow, 1).Value _
            And newWs.Cells(newRow - 1, 6).Value = newWs.Cells(newRow, 6).Value _
        Then
            newWs.Cells(newRow, 10).Value = newWs.Cells(newRow, 9).Value - newWs.Cells(newRow - 1, 9).Value
            If newWs.Cells(newRow, 10).Value < 60 Then
                nErrorCnt = nErrorCnt + 1
                newWs.Rows(newRow).Interior.color = RGB(255, 0, 0)
            End If
        End If
    Next newRow
    
    newWs.Cells(2, 12).Value = "시간오류"
    newWs.Cells(2, 13).Value = nErrorCnt

    
    '=================================================
    ' 6단계 : 각 노선별로 첫차 막자 자동 계산
    '=================================================
     
     SearchFirstBusLastBus newWs
    
    
    newWs.Columns.AutoFit
    
    
End Sub
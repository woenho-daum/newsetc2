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
            c.Value = Replace(c.Value, Chr(10), " ")
        End If
    Next c

End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectItem 
   Caption         =   "선택화면"
   ClientHeight    =   8172
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10164
   OleObjectBlob   =   "frmSelectItem.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmSelectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SelectSheet(xlWB As Object) As Object

    Dim i As Long
    Dim itm As ListItem

    Me.Tag = ""
    Me.frmTitle.Caption = "참조할 시트를 선택하세요."

    With Me.lvItems

        .ListItems.Clear
        .ColumnHeaders.Clear

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False

        .ColumnHeaders.Add , , "시트명", 180
        .ColumnHeaders.Add , , "번호", 50

        For i = 1 To xlWB.Sheets.Count

            Set itm = .ListItems.Add(, , xlWB.Sheets(i).Name)
            itm.SubItems(1) = CStr(i)

            '선택 결과 저장
            itm.Tag = xlWB.Sheets(i).Name

        Next i

    End With

    Me.Show vbModal

    If Me.Tag = "" Then
        Set SelectSheet = Nothing
    Else
        Set SelectSheet = xlWB.Sheets(Me.Tag)
    End If

End Function


Private Sub UserForm_Initialize()

    Me.Tag = ""

    'Me.KeyPreview = True    '이거 안키면 엔터키로 그냥 화면 닫힌다
 
    With Me.lvItems
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideSelection = False
        .LabelEdit = lvwManual      '라벨 수정 금지
    End With

End Sub

Private Sub UserForm_Activate()
    If lvItems.ListItems.Count > 0 Then
        lvItems.SetFocus
        lvItems.ListItems(1).Selected = True
        lvItems.ListItems(1).EnsureVisible
    End If
End Sub

Private Sub btnOK_Click()

    If lvItems.SelectedItem Is Nothing Then
        MsgBox "항목을 선택하세요.", vbExclamation
        Exit Sub
    End If

    Me.Tag = lvItems.SelectedItem.Tag
    Me.Hide

End Sub


Private Sub btnCancel_Click()
    ' ESC 누르면 그냥 무조건 이리온다
    Me.Tag = ""
    Me.Hide

End Sub

Private Sub lvItems_DblClick()
    btnOK_Click
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

    ' default = true 로 설정하면 이 함수 안탄다, 엔터키시에 그냥 default=true 탄다
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        btnOK_Click
    End If
    
End Sub




VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChatBotForm 
   Caption         =   "함수 자동 생성기"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10605
   OleObjectBlob   =   "ChatBotForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "ChatBotForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================
' 엑셀 함수 생성 챗봇 폼 코드
' =============================

Private Sub btnSend_Click()
    Dim question As String
    Dim answer As String
    
    question = Trim(Me.txtQuestion.text)
    If question = "" Then
        MsgBox "질문을 입력하세요!", vbExclamation
        Me.txtQuestion.SetFocus
        Exit Sub
    End If
    
    Me.lblStatus.Caption = "상태: 서버에 요청 중..."
    Me.Repaint
    
    answer = AskExcelFormula(question)
    
    If answer <> "" Then
        Me.txtAnswer.text = answer
        Me.lblStatus.Caption = "상태: 답변 완료"
    Else
        Me.txtAnswer.text = "응답이 없습니다."
        Me.lblStatus.Caption = "상태: 오류 발생"
    End If
End Sub

Private Sub lblStatus_Click()

End Sub

Private Sub txtAnswer_Change()

End Sub

Private Sub txtQuestion_Change()
    ' 텍스트 입력 변경 시 필요한 동작이 있다면 여기에 추가
End Sub

Private Sub txtQuestion_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnSend_Click
        KeyCode = 0 ' 줄바꿈 방지 (Enter 기본 동작 막기)
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.txtQuestion.text = "엑셀 함수/수식에 대해 질문하세요."
    Me.txtAnswer.text = "여기에 답변(엑셀 함수/수식)이 표시됩니다."
    Me.lblStatus.Caption = "상태: 대기 중"
End Sub


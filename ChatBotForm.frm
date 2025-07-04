VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChatBotForm 
   Caption         =   "�Լ� �ڵ� ������"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10605
   OleObjectBlob   =   "ChatBotForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "ChatBotForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================
' ���� �Լ� ���� ê�� �� �ڵ�
' =============================

Private Sub btnSend_Click()
    Dim question As String
    Dim answer As String
    
    question = Trim(Me.txtQuestion.text)
    If question = "" Then
        MsgBox "������ �Է��ϼ���!", vbExclamation
        Me.txtQuestion.SetFocus
        Exit Sub
    End If
    
    Me.lblStatus.Caption = "����: ������ ��û ��..."
    Me.Repaint
    
    answer = AskExcelFormula(question)
    
    If answer <> "" Then
        Me.txtAnswer.text = answer
        Me.lblStatus.Caption = "����: �亯 �Ϸ�"
    Else
        Me.txtAnswer.text = "������ �����ϴ�."
        Me.lblStatus.Caption = "����: ���� �߻�"
    End If
End Sub

Private Sub lblStatus_Click()

End Sub

Private Sub txtAnswer_Change()

End Sub

Private Sub txtQuestion_Change()
    ' �ؽ�Ʈ �Է� ���� �� �ʿ��� ������ �ִٸ� ���⿡ �߰�
End Sub

Private Sub txtQuestion_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnSend_Click
        KeyCode = 0 ' �ٹٲ� ���� (Enter �⺻ ���� ����)
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.txtQuestion.text = "���� �Լ�/���Ŀ� ���� �����ϼ���."
    Me.txtAnswer.text = "���⿡ �亯(���� �Լ�/����)�� ǥ�õ˴ϴ�."
    Me.lblStatus.Caption = "����: ��� ��"
End Sub


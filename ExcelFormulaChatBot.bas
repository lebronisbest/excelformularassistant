Attribute VB_Name = "Module1"
Option Explicit

' =====================================
' Excel Formula ChatBot (엑셀 첨부 전용)
' =====================================

Private Const SERVER_URL As String = "http://127.0.0.1:8000"

' 임시 파일 경로 반환
Public Function GetTempExcelPath() As String
    GetTempExcelPath = Environ$("TEMP") & "\chatgpt_temp_upload.xlsx"
End Function

' 현재 워크북을 임시 파일로 저장
Public Sub SaveWorkbookToTemp()
Attribute SaveWorkbookToTemp.VB_ProcData.VB_Invoke_Func = "C\n14"
    ThisWorkbook.SaveCopyAs GetTempExcelPath()
End Sub

' 엑셀 파일과 질문을 함께 전송 (폼에서 이 함수만 호출)
Public Function AskExcelFormula(question As String) As String
    Dim http As Object
    Dim tempFile As String
    Dim boundary As String
    Dim response As String
    Dim fileStream As Object
    Dim bodyStream As Object
    Dim header As String, fileHeader As String, footer As String

    tempFile = GetTempExcelPath()
    SaveWorkbookToTemp

    boundary = "----WebKitFormBoundary" & Hex(Timer * 1000000)
    Set http = CreateObject("MSXML2.XMLHTTP")

    header = "--" & boundary & vbCrLf
    header = header & "Content-Disposition: form-data; name=""question""" & vbCrLf & vbCrLf

    fileHeader = vbCrLf & "--" & boundary & vbCrLf
    fileHeader = fileHeader & "Content-Disposition: form-data; name=""file""; filename=""chatgpt_temp_upload.xlsx""" & vbCrLf
    fileHeader = fileHeader & "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" & vbCrLf & vbCrLf

    footer = vbCrLf & "--" & boundary & "--" & vbCrLf

    Set bodyStream = CreateObject("ADODB.Stream")
    bodyStream.Type = 1 ' Binary
    bodyStream.Open

    Dim headerBytes() As Byte
    headerBytes = StrConv(header, vbFromUnicode)
    bodyStream.Write headerBytes

    Dim questionStream As Object
    Set questionStream = CreateObject("ADODB.Stream")
    questionStream.Type = 2 ' Text
    questionStream.Charset = "utf-8"
    questionStream.Open
    questionStream.WriteText question
    questionStream.Position = 0
    questionStream.Type = 1 ' Binary로 변경
    bodyStream.Write questionStream.Read
    questionStream.Close

    Dim fileHeaderBytes() As Byte
    fileHeaderBytes = StrConv(fileHeader, vbFromUnicode)
    bodyStream.Write fileHeaderBytes

    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1
    fileStream.Open
    fileStream.LoadFromFile tempFile
    bodyStream.Write fileStream.Read
    fileStream.Close

    Dim footerBytes() As Byte
    footerBytes = StrConv(footer, vbFromUnicode)
    bodyStream.Write footerBytes

    bodyStream.Position = 0

    http.Open "POST", SERVER_URL & "/upload_chat", False
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.send bodyStream.Read

    If http.Status = 200 Then
        response = http.responseText
        AskExcelFormula = ExtractFormula(response)
    Else
        AskExcelFormula = "HTTP 오류: " & http.Status
    End If

    bodyStream.Close
    Set http = Nothing
End Function

' JSON에서 답변(엑셀 함수/수식)만 추출 (JsonConverter 사용)
Private Function ExtractFormula(jsonText As String) As String
    Dim json As Object
    On Error GoTo ErrHandler
    Set json = JsonConverter.ParseJson(jsonText)
    ExtractFormula = ""
    If Not json Is Nothing Then
        If json.Exists("answer") Then
            ExtractFormula = json.Item("answer")
        End If
    End If
    Exit Function
ErrHandler:
    ExtractFormula = ""
End Function

Sub ShowExcelFormulaChatBot()
    ChatBotForm.Show
End Sub


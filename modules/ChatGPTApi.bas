Attribute VB_Name = "ChatGPTApi"
Option Compare Database
Option Explicit

Private Const CHATGPT_KEY = ""

Function GetChatGPTApi(ByVal StrMessage As String) As Dictionary
    Dim webJSON As New webJSON
    Dim webClient As New webClient
    Dim webRequest As New webRequest
    Dim webResponse As webResponse
    Dim DicMes1 As New Dictionary
    Dim root As Dictionary

    webClient.BaseUrl = "https://api.openai.com/v1/chat/completions"

    webRequest.Resource = ""
    webRequest.Method = WebMethod.Httppost
    webRequest.Format = WebFormat.Json

    webRequest.AddHeader "Authorization", "Bearer " & CHATGPT_KEY

    webRequest.AddBodyParameter "model", "gpt-5"

    DicMes1.Add "role", "user"

    DicMes1.Add "content", StrMessage

    webRequest.AddBodyParameter "messages", Array(DicMes1)

    webRequest.AddBodyParameter "temperature", 1

    webRequest.AddBodyParameter "max_completion_tokens", Null

    webRequest.AddBodyParameter "top_p", 1

    webRequest.AddBodyParameter "frequency_penalty", 0

    webRequest.AddBodyParameter "presence_penalty", 0

    Set webResponse = webClient.Execute(webRequest)

    If webResponse.StatusCode <> WebStatusCode.Ok Then
        MsgBox (webResponse.Content)
        Exit Function
    End If

    Set root = webResponse.Data()


    Debug.Print webClient.BaseUrl & webRequest.FormattedResource
    Debug.Print webResponse.StatusCode, webResponse.StatusDescription
    Debug.Print webResponse.Content

    If Not root Is Nothing Then

        'Debug.Print "data", root("data")
       ' Debug.Print "pagination", root("pagination")
     '   Debug.Print "from_symbol", root("from_symbol")

    End If

    Set GetChatGPTApi = root

End Function

Function Uni2Utf(ByVal Text As String) As String
  Dim v As Long
  Dim i As Long

  For i = 1 To Len(Text)
    v = AscW(Mid(Text, i, 1))
    Select Case v
      Case Is < 128
        Uni2Utf = Uni2Utf & Mid(Text, i, 1)
      Case Is < 2048
        Uni2Utf = Uni2Utf & Chr(((v And 1984) / 64) Or 192)
        Uni2Utf = Uni2Utf & Chr((v And 63) Or 128)
      Case Else
        Uni2Utf = Uni2Utf & Chr(((v And 61440) / 4096) Or 224)
        Uni2Utf = Uni2Utf & Chr(((v And 4032) / 64) Or 128)
        Uni2Utf = Uni2Utf & Chr((v And 63) Or 128)
    End Select
  Next
End Function

Function Utf2Uni(ByVal Text As String) As String
  Dim v As Long
  Dim i As Long: i = 1

  Do While i <= Len(Text)
    v = Asc(Mid(Text, i, 1))
    Select Case v
      Case Is < 128
        Utf2Uni = Utf2Uni & Mid(Text, i, 1)
        i = i + 1
      Case Is < 224
        Utf2Uni = Utf2Uni & ChrW((v And 63) * 64 + (Asc(Mid(Text, i + 1, 1)) And 63))
        i = i + 2
      Case Else
        Utf2Uni = Utf2Uni & ChrW(((v And 31) * 4096) + _
          ((Asc(Mid(Text, i + 1, 1)) And 63) * 64) + _
          (Asc(Mid(Text, i + 2, 1)) And 63))
        i = i + 3
    End Select
  Loop
End Function

Attribute VB_Name = "LanguageApi"
Option Compare Database
Option Explicit

Private Const GOOGLEAPI_KEY = ""

Function GetLanguageApi(ByVal StrQ As String, ByVal FromLanguage As String, ByVal ToLanguage As String) As Dictionary
    Dim webClient As New webClient
    Dim webRequest As New webRequest
    Dim webResponse As webResponse
    Dim root As Dictionary

    webClient.BaseUrl = "https://translation.googleapis.com/language/translate/v2"

    webRequest.Resource = ""
    webRequest.Method = WebMethod.Httppost
    webRequest.Format = WebFormat.Json
    With webRequest
        .AddQuerystringParam "q", StrQ
        .AddQuerystringParam "target", ToLanguage
        .AddQuerystringParam "format", "text"
        .AddQuerystringParam "source", FromLanguage
        .AddQuerystringParam "model", "base"
        .AddQuerystringParam "key", GOOGLEAPI_KEY
    End With

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

        Debug.Print "currency_updated", root("currency_updated")
        Debug.Print "from_amount", root("from_amount")
        Debug.Print "from_code", root("from_code")
        Debug.Print "from_name", root("from_name")
        Debug.Print "from_symbol", root("from_symbol")
        Debug.Print "to_amount", root("to_amount")
        Debug.Print "to_code", root("to_code")
        Debug.Print "to_name", root("to_name")
        Debug.Print "to_symbol", root("to_symbol")

    End If

    Set GetLanguageApi = root

End Function

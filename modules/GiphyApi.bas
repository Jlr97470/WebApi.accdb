Attribute VB_Name = "GiphyApi"
Option Compare Database
Option Explicit


Private Const GIPHY_KEY = ""

Function GetGiphyApi(ByVal StrBundle As String) As Dictionary
    Dim webClient As New webClient
    Dim webRequest As New webRequest
    Dim webResponse As webResponse
    Dim root As Dictionary

    webClient.BaseUrl = "https://api.giphy.com/v1/stickers/search"

    webRequest.Resource = ""
    webRequest.Method = WebMethod.Httpget
    webRequest.Format = WebFormat.json
    With webRequest
        .AddQuerystringParam "api_key", GIPHY_KEY
        .AddQuerystringParam "bundle", "sticker_layering"
        .AddQuerystringParam "q", StrBundle
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

        'Debug.Print "data", root("data")
       ' Debug.Print "pagination", root("pagination")
        Debug.Print "from_symbol", root("from_symbol")

    End If

    Set GetGiphyApi = root

End Function

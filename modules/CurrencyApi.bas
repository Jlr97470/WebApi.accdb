Attribute VB_Name = "CurrencyApi"
Option Compare Database
Option Explicit


Function GetConvertCurrencyApi(ByVal FromCurrency As String, ByVal ToCurrency As String, ByVal CurAmount As Currency) As Dictionary
    Dim webClient As New webClient
    Dim webRequest As New webRequest
    Dim webResponse As webResponse
    Dim root As Dictionary

    webClient.BaseUrl = "http://www.geoplugin.net/currency/json.gp"

    webRequest.Resource = ""
    webRequest.Method = WebMethod.Httpget
    webRequest.Format = WebFormat.Json
    With webRequest
        .AddQuerystringParam "from", FromCurrency
        .AddQuerystringParam "to", ToCurrency
        .AddQuerystringParam "amount", CurAmount
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

    Set GetConvertCurrencyApi = root

End Function

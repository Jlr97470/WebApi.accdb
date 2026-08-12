Attribute VB_Name = "ActorApi"
Option Compare Database
Option Explicit


Function GetActorApi(ByVal StrAction As String, ByRef StrName() As String, ByRef StrData() As String) As Dictionary
    Dim webClient As New webClient
    Dim webRequest As New webRequest
    Dim webResponse As webResponse
    Dim StrColName As Variant
    Dim StrColData As Variant
    Dim Index As Integer
    Dim root As Dictionary

    webClient.BaseUrl = "https://localhost/PHP/Actor.php"

    webRequest.Resource = ""
    webRequest.Method = WebMethod.Httppost
    webRequest.Format = WebFormat.FormUrlEncoded
    webRequest.ResponseFormat = json

    With webRequest
        .AddBodyParameter "action", StrAction

        If StrName(0) <> "" Then

            For Index = LBound(StrName) To UBound(StrName)

                StrColName = StrName(Index)

                .AddBodyParameter "name" & Index, StrColName

            Next

        End If

        If StrData(0) <> "" Then

            For Index = LBound(StrData) To UBound(StrData)

                StrColData = StrData(Index)

                .AddBodyParameter "data" & Index, StrColData

            Next

        End If
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

     '   Debug.Print "currency_updated", root("currency_updated")
     '   Debug.Print "from_amount", root("from_amount")
     '   Debug.Print "from_code", root("from_code")
     '   Debug.Print "from_name", root("from_name")
     '   Debug.Print "from_symbol", root("from_symbol")
     '   Debug.Print "to_amount", root("to_amount")
     '   Debug.Print "to_code", root("to_code")
     '   Debug.Print "to_name", root("to_name")
     '  Debug.Print "to_symbol", root("to_symbol")

    End If

    Set GetActorApi = root

End Function

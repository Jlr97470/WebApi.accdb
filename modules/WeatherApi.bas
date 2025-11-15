Attribute VB_Name = "WeatherApi"
Option Compare Database
Option Explicit

Function GetWeatherFromApi(ByVal latitude As String, ByVal longitude As String) As Dictionary
    Dim webClient As New webClient
    Dim webRequest As New webRequest
    Dim webResponse As webResponse
    Dim root As Dictionary
    Dim weather As Dictionary

    webClient.BaseUrl = "https://api.open-meteo.com/v1/"

    webRequest.Resource = "forecast"
    webRequest.Method = WebMethod.Httpget
    webRequest.Format = WebFormat.Json
    With webRequest
        .AddQuerystringParam "latitude", latitude
        .AddQuerystringParam "longitude", longitude
        .AddQuerystringParam "current_weather", "true"
        .AddQuerystringParam "temperature_unit", "celsius"
        .AddQuerystringParam "windspeed_unit", "kmh"
        .AddQuerystringParam "current", "is_day"
    End With

    Set webResponse = webClient.Execute(webRequest)

    If webResponse.StatusCode <> WebStatusCode.Ok Then
        MsgBox (webResponse.Content)
        Exit Function
    End If

    Set root = webResponse.Data()
    Set weather = root("current_weather")

    Debug.Print webClient.BaseUrl & webRequest.FormattedResource
    Debug.Print webResponse.StatusCode, webResponse.StatusDescription
    Debug.Print webResponse.Content

    If Not weather Is Nothing Then

        Debug.Print "time", weather("time"), root("timezone")
        Debug.Print "temperature", weather("temperature"), "Degrees C"
        Debug.Print "wind speed", weather("windspeed"), "KMH"
        Debug.Print "weather code", weather("weathercode"), ""

    End If

    Set GetWeatherFromApi = root

End Function

Sub GetLocalWeather()
    Debug.Print "-------" & vbCrLf & "VBA-Web" & vbCrLf & "-------"
    Call GetWeatherFromApi("39.37", "-104.86")
End Sub

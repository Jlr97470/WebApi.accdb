Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnGetWeather_Click()
    Dim db As DAO.Database
    Dim RstCity As DAO.Recordset
    Dim DicResult As Dictionary

    If IsNull(Me.lstCities.Value) Then
        MsgBox "Please select a city"
        Exit Sub
    End If

    Me.txtTemperature.Value = ""
    Me.txtWindSpeed.Value = ""
    Me.txtWindDirection.Value = ""
    Me.lblStatus.Caption = "Getting weather..."
    DoEvents


    Set db = CurrentDb
    strSQL = "SELECT Latitude, Longitude FROM CityLatLong WHERE ID = " & Me.lstCities.Value
    Set RstCity = db.OpenRecordset(strSQL, dbOpenSnapshot)


    Set DicResult = GetWeatherFromApi(RstCity!latitude, RstCity!longitude)

    RstCity.Close

    If Not DicResult Is Nothing Then

        Me.txtTemperature.Value = DicResult("current_weather")("temperature")
        Me.txtWindSpeed.Value = DicResult("current_weather")("windspeed")
        Me.txtWindDirection.Value = DicResult("current_weather")("winddirection")
        Me.txtweathercode.Value = DicResult("current_weather")("weathercode")

    End If

    Me.lblStatus.Caption = "Done"

    Set db = Nothing
    Set RstCity = Nothing
    Set DicResult = Nothing

End Sub

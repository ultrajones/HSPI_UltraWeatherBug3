Imports System.Net
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class hspi_weatherbug_api

  Private _querySuccess As ULong = 0
  Private _queryFailure As ULong = 0

  Public Sub New()

  End Sub

  Public Function QuerySuccessCount() As ULong
    Return _querySuccess
  End Function

  Public Function QueryFailureCount() As ULong
    Return _queryFailure
  End Function

  ''' <summary>
  ''' Determines if the API key is valid
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidAPIKey() As Boolean

    Try

      If gAPIKeyPrimary.Length = 0 Then
        Return False
      ElseIf gAPIKeySecondary.Length = 0 Then
        Return True
      Else
        Return True
      End If

    Catch pEx As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Gets the WeatherBug Locations
  ''' </summary>
  ''' <param name="searchString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetLocations(searchString As String) As ArrayList

    Try
      Dim results As New ArrayList
      Dim strURL As String = String.Format("https://earthnetworks.azure-api.net/getLocations/data/locations/v2/location?searchString={0}&subscription-key={1}", searchString, gAPIKeyPrimary)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          Dim locations As List(Of Location) = js.Deserialize(Of List(Of Location))(JSONString)

          For Each Location As Location In locations
            Dim MyLocation As New Specialized.StringDictionary

            Dim la As Double = 0.0
            Double.TryParse(Location.la, la)

            Dim lo As Double = 0.0
            Double.TryParse(Location.lo, lo)

            MyLocation.Add("CityId", Location.ci)
            MyLocation.Add("CityName", Location.cn)
            MyLocation.Add("Territory", Location.t)
            MyLocation.Add("StateCode", Location.sc)
            MyLocation.Add("Country", Location.c)
            MyLocation.Add("Lattitude", la.ToString(CultureInfo.InvariantCulture))
            MyLocation.Add("Longitude", lo.ToString(CultureInfo.InvariantCulture))
            MyLocation.Add("Dma", Location.d)
            MyLocation.Add("Zip", Location.z)

            results.Add(MyLocation)
          Next

        End Using

      End Using

      _querySuccess += 1

      Return results

    Catch pEx As Exception
      '
      ' Process the error
      '
      ProcessError(pEx, "GetLocations")

      _queryFailure += 1
    End Try

    Return New ArrayList

  End Function

  ''' <summary>
  ''' Gets the WeatherBug Locations
  ''' </summary>
  ''' <param name="searchString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStationList(searchString As String) As ArrayList

    Try
      Dim results As New ArrayList

      Dim strURL As String = String.Format("https://earthnetworks.azure-api.net/getStationList/data/locations/v3/stationlist?location={0}&locationtype=latitudelongitude&subscription-key={1}", searchString, gAPIKeyPrimary)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          Dim StationList As StationList = js.Deserialize(Of StationList)(JSONString)

          For Each Station As Station In StationList.r.s

            Dim MyStation As New Specialized.StringDictionary

            Dim la As Double = 0.0
            Double.TryParse(Station.la, la)

            Dim lo As Double = 0.0
            Double.TryParse(Station.lo, lo)

            MyStation.Add("StationId", Station.si)
            MyStation.Add("ProviderId", Station.pi)
            MyStation.Add("ProviderName", Station.pn)
            MyStation.Add("StationName", Station.sn)
            MyStation.Add("Lattitude", la.ToString(CultureInfo.InvariantCulture))
            MyStation.Add("Longitude", lo.ToString(CultureInfo.InvariantCulture))
            MyStation.Add("EASL", Station.easl)
            MyStation.Add("Distance", Station.d)
            MyStation.Add("Unit", "miles")

            results.Add(MyStation)
          Next

        End Using

      End Using

      _querySuccess += 1

      Return results

    Catch pEx As Exception
      '
      ' Process the error
      '
      ProcessError(pEx, "GetStationList")

      _queryFailure += 1
    End Try

    Return New ArrayList

  End Function

  ''' <summary>
  ''' Gets the Realtime Weather from WeatherBug
  ''' </summary>
  ''' <param name="searchString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetRealtimeWeather(searchString As String) As RealTimeWeather

    Dim RealTimeWeather As New RealTimeWeather

    Try

      Dim unitType As String = hs.GetINISetting("Options", "UnitType", "0", gINIFile)
      Dim units As String = IIf(unitType = "1", "metric", "english")

      Dim strURL As String = String.Format("https://earthnetworks.azure-api.net/data/observations/v1/current?location={0}&locationtype=latitudelongitude&units={1}&verbose=true&subscription-key={2}", searchString, units, gAPIKeyPrimary)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          RealTimeWeather = js.Deserialize(Of RealTimeWeather)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _queryFailure += 1

    End Try

    Return RealTimeWeather

  End Function

  ''' <summary>
  ''' Gets the Forecast from WeatherBug
  ''' </summary>
  ''' <param name="searchString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWeatherForecast(searchString As String) As WeatherForecast

    Dim WeatherForecast As New WeatherForecast

    Try

      Dim unitType As String = hs.GetINISetting("Options", "UnitType", "0", gINIFile)
      Dim units As String = IIf(unitType = "1", "metric", "english")

      Dim strURL As String = String.Format("https://earthnetworks.azure-api.net/data/forecasts/v1/daily?location={0}&locationtype=latitudelongitude&units={1}&verbose=true&subscription-key={2}", searchString, units, gAPIKeyPrimary)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd().Replace(": true", ": ""true""").Replace(": false", ": ""false""")

          Dim js As New JavaScriptSerializer()
          WeatherForecast = js.Deserialize(Of WeatherForecast)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _queryFailure += 1

    End Try

    Return WeatherForecast

  End Function

  ''' <summary>
  ''' Gets the Alerts from WeatherBug
  ''' </summary>
  ''' <param name="searchString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWeatherAlerts(searchString As String) As WeatherAlerts

    Dim WeatherAlerts As New WeatherAlerts

    Try

      Dim strURL As String = String.Format("https://earthnetworks.azure-api.net/getPublishedAlerts/data/alerts/v1/alerts?location={0}&locationtype=latitudelongitude&verbose=true&subscription-key={1}", searchString, gAPIKeyPrimary)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          WeatherAlerts = js.Deserialize(Of WeatherAlerts)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _queryFailure += 1
    End Try

    Return WeatherAlerts

  End Function

  Public Function GetAlertSignificance(alertType As String) As String

    Dim Significance As String = String.Empty

    Try

      Significance = Regex.Match(alertType, "[A-Z]$").ToString
      Select Case Significance
        Case "W" : Return "Warning"
        Case "F" : Return "Forecast"
        Case "A" : Return "Watch"
        Case "O" : Return "Outlook"
        Case "Y" : Return "Advisory"
        Case "N" : Return "Synopsis"
        Case "S" : Return "Statement"
      End Select

    Catch pEx As Exception

    End Try

    Return alertType

  End Function

  Public Function GetAlertPhenomena(alertType As String) As String

    Dim Phenomena As String = String.Empty

    Try

      Phenomena = Regex.Match(alertType, "^[A-Z][A-Z]").ToString

      Select Case Phenomena
        Case "AF" : Return "Ashfall "
        Case "AS" : Return "Air Stagnation"
        Case "BS" : Return "Blowing Snow"
        Case "BW" : Return "Brisk Wind"
        Case "BZ" : Return "Blizzard"
        Case "CF" : Return "Coastal Flood"
        Case "DS" : Return "Dust Storm"
        Case "DU" : Return "Blowing Dust"
        Case "EC" : Return "Extreme Cold"
        Case "EH" : Return "Excessive Heat"
        Case "EW" : Return "Extreme Wind"
        Case "FA" : Return "Areal Flood"
        Case "FF" : Return "Flash Flood"
        Case "FG" : Return "Dense Fog"
        Case "FL" : Return "Flood"
        Case "FR" : Return "Frost"
        Case "FW" : Return "Fire Weather"
        Case "FZ" : Return "Freeze"
        Case "GL" : Return "Gale"
        Case "HF" : Return "Hurricane Force Wind"
        Case "HI" : Return "Inland Hurricane"
        Case "HS" : Return "Heavy Snow"
        Case "HT" : Return "Heat"
        Case "HU" : Return "Hurricane"
        Case "HW" : Return "High Wind"
        Case "HY" : Return "Hydrologic"
        Case "HZ" : Return "Hard Freeze"
        Case "IP" : Return "Sleet"
        Case "IS" : Return "Ice Storm"
        Case "LB" : Return "Lake Effect Snow and Blowing Snow"
        Case "LE" : Return "Lake Effect Snow"
        Case "LO" : Return "Low Water"
        Case "LS" : Return "Lakeshore Flood"
        Case "LW" : Return "Lake Wind"
        Case "MA" : Return "Marine"
        Case "RB" : Return "Small Craft for Rough Bar"
        Case "SB" : Return "Snow and Blowing Snow"
        Case "SC" : Return "Small Craft"
        Case "SE" : Return "Hazardous Seas"
        Case "SI" : Return "Small Craft for Winds"
        Case "SM" : Return "Dense Smoke"
        Case "SN" : Return "Snow"
        Case "SR" : Return "Storm"
        Case "SU" : Return "High Surf"
        Case "SV" : Return "Severe Thunderstorm"
        Case "SW" : Return "Small Craft for Hazardous Seas"
        Case "TI" : Return "Inland Tropical Storm"
        Case "TO" : Return "Tornado"
        Case "TR" : Return "Tropical Storm"
        Case "TS" : Return "Tsunami TY Typhoon"
        Case "UP" : Return "Ice Accretion"
        Case "WC" : Return "Wind Chill"
        Case "WI" : Return "Wind"
        Case "WS" : Return "Winter Storm"
        Case "WW" : Return "Winter Weather"
        Case "ZF" : Return "Freezing Fog"
        Case "ZR" : Return "Freezing Rain"
      End Select

    Catch pEx As Exception

    End Try

    Return alertType

  End Function

  Public Function GetWeatherIcons(searchString As String) As WeatherIcons

    Dim WeatherIcons As New WeatherIcons

    Try

      Dim unitType As String = hs.GetINISetting("Options", "UnitType", "0", gINIFile)
      Dim units As String = IIf(unitType = "1", "metric", "english")

      Dim strURL As String = String.Format("https://earthnetworks.azure-api.net/getSkyConditionIcons/resources/v1/icons?IconSet={0}&subscription-key={1}", searchString, gAPIKeyPrimary)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          WeatherIcons = js.Deserialize(Of WeatherIcons)(JSONString)

          Dim Width As Integer = 0
          Dim Height As Integer = 0

          Integer.TryParse(WeatherIcons.Width, Width)
          Integer.TryParse(WeatherIcons.Height, Height)

          For Each WeatherIcon As WeatherIcon In WeatherIcons.Icons

          Next

        End Using

      End Using

    Catch pEx As Exception

    End Try

    Return WeatherIcons

  End Function

  Private Sub Util()

    Dim strResult As String = String.Empty

    For i As Integer = 0 To 282
      Dim strPerdiction As String = GetWeatherConditionNameById(i)
      strPerdiction = Regex.Replace(strPerdiction, "\d\d% Change of", "", RegexOptions.IgnoreCase)
      strPerdiction = Regex.Replace(strPerdiction, "Change of (Heavy|Light) ", "", RegexOptions.IgnoreCase)
      strPerdiction = Regex.Replace(strPerdiction, "Change of ", "", RegexOptions.IgnoreCase)

      Dim result As Integer = GetWeatherConditionValueByRegex(strPerdiction)

      strResult &= String.Format("Case ""{0}"" : Return {1} ' {2}", i.ToString.PadLeft(3, "0"), result.ToString, strPerdiction) & vbCrLf
    Next

  End Sub

#Region "WeatherBug Location"

  '  [
  '  {
  '    "ci": "US36E0032",
  '    "cn": "East Rochester",
  '    "t": "New York",
  '    "sc": "NY",
  '    "c": "United States",
  '    "la": 43.1115,
  '    "lo": -77.4902,
  '    "d": "538",
  '    "z": "14445"
  '  }
  ']

  <Serializable()> _
  Public Class Location
    Public Property ci As String = String.Empty
    Public Property cn As String = String.Empty
    Public Property t As String = String.Empty
    Public Property sc As String = String.Empty
    Public Property c As String = String.Empty
    Public Property la As String = String.Empty ' Double
    Public Property lo As String = String.Empty ' Double
    Public Property d As String = String.Empty
    Public Property z As String = String.Empty
  End Class

#End Region

#Region "WeatherBug Station List"

  '  {
  '  "i": "27721135-5a4e-481d-bcd9-774c86b25a3a",
  '  "c": 200,
  '  "e": null,
  '  "r": {
  '    "s": [
  '      {
  '        "si": "CHLBY",
  '        "pi": 3,
  '        "pn": "Earth Networks Inc",
  '        "sn": "Shadbush Center School",
  '        "la": 42.6483333333333,
  '        "lo": -83.0638888888889,
  '        "easl": null,
  '        "df": null,
  '        "d": 10.76
  '      }
  ']

  <Serializable()> _
  Private Class StationList
    Public Property i As String = String.Empty ' id
    Public Property c As String = String.Empty ' Code
    Public Property e As String = String.Empty ' Error Message  ' Integer
    Public Property r As Result
  End Class

  <Serializable()> _
  Private Class Result
    Public Property s As List(Of Station)
  End Class

  <Serializable()> _
  Private Class Station
    Public Property si As String = String.Empty   ' Station Id
    Public Property pi As String = String.Empty   ' Provider Id
    Public Property pn As String = String.Empty   ' Provider Name
    Public Property sn As String = String.Empty   ' Station Name
    Public Property la As String = String.Empty   ' Lattitude   ' Double
    Public Property lo As String = String.Empty   ' Logitude    ' Double
    Public Property easl As String = String.Empty ' Elivation Above Sea Level
    Public Property df As String = String.Empty   ' Display Flag
    Public Property d As String = String.Empty    ' Distance
  End Class

#End Region

#Region "WeatherBug Real Time Weather"

  '{
  '  "key": null,
  '  "stationId": "UTCPP",
  '  "providerId": 3,
  '  "observationTimeLocalStr": "2014-04-07T19:54:00",
  '  "observationTimeUtcStr": "2014-04-07T23:54:00",
  '  "iconCode": null,
  '  "altimeter": null,
  '  "altimeterRate": null,
  '  "dewPoint": 4.6,
  '  "dewPointRate": null,
  '  "heatIndex": 5.4,
  '  "humidity": 95.0,
  '  "humidityRate": 7.2,
  '  "pressureSeaLevel": 998.0,
  '  "pressureSeaLevelRate": -1.0,
  '  "rainDaily": 0.0,
  '  "rainRate": 0.0,
  '  "rainMonthly": 0.0,
  '  "rainYearly": 0.0,
  '  "snowDaily": null,
  '  "snowRate": null,
  '  "snowMonthly": null,
  '  "snowYearly": null,
  '  "temperature": 5.4,
  '  "temperatureRate": -1.1,
  '  "visibility": null,
  '  "visibilityRate": null,
  '  "windChill": 5.4,
  '  "windSpeed": 1.1,
  '  "windDirection": 7,
  '  "windSpeedAvg": 3.9,
  '  "windDirectionAvg": 8,
  '  "windGustHourly": 14.1,
  '  "windGustTimeLocalHourlyStr": "2014-04-07T19:27:00",
  '  "windGustTimeUtcHourlyStr": "2014-04-07T23:27:00",
  '  "windGustDirectionHourly": 32,
  '  "windGustDaily": 21.2,
  '  "windGustTimeLocalDailyStr": "2014-04-07T12:30:00",
  '  "windGustTimeUtcDailyStr": "2014-04-07T16:30:00",
  '  "windGustDirectionDaily": 65,
  '  "observationTimeAdjustedLocalStr": "2014-04-07T19:54:51",
  '  "feelsLike": 5.4
  '}

  <Serializable()> _
  Public Class RealTimeWeather
    Public Property key As String = String.Empty
    Public Property stationId As String = String.Empty
    Public Property providerId As String = String.Empty
    Public Property observationTimeLocalStr As String = String.Empty
    Public Property observationTimeUtcStr As String = String.Empty
    Public Property iconCode As String = String.Empty
    Public Property altimeter As String = String.Empty
    Public Property altimeterRate As String = String.Empty
    Public Property dewPoint As String = String.Empty ' Double
    Public Property dewPointRate As String = String.Empty ' Double
    Public Property heatIndex As String = String.Empty ' Double
    Public Property humidity As String = String.Empty ' Double
    Public Property humidityRate As String = String.Empty ' Double
    Public Property pressureSeaLevel As String = String.Empty ' Double
    Public Property pressureSeaLevelRate As String = String.Empty ' Double
    Public Property rainDaily As String = String.Empty ' Double
    Public Property rainRate As String = String.Empty ' Double
    Public Property rainMonthly As String = String.Empty ' Double
    Public Property rainYearly As String = String.Empty ' Double
    Public Property snowDaily As String = String.Empty ' Double
    Public Property snowRate As String = String.Empty ' Double
    Public Property snowMonthly As String = String.Empty ' Double
    Public Property snowYearly As String = String.Empty ' Double
    Public Property temperature As String = String.Empty ' Double
    Public Property temperatureRate As String = String.Empty ' Double
    Public Property visibility As String = String.Empty ' Double
    Public Property visibilityRate As String = String.Empty ' Double
    Public Property windChill As String = String.Empty ' Double
    Public Property windSpeed As String = String.Empty ' Double
    Public Property windDirection As String = String.Empty ' Integer
    Public Property windSpeedAvg As String = String.Empty ' Double
    Public Property windDirectionAvg As String = String.Empty ' Integer
    Public Property windGustHourly As String = String.Empty ' Double
    Public Property windGustTimeLocalHourlyStr As String = String.Empty
    Public Property windGustTimeUtcHourlyStr As String = String.Empty
    Public Property windGustDirectionHourly As String = String.Empty ' Integer
    Public Property windGustDaily As String = String.Empty ' Double
    Public Property windGustTimeLocalDailyStr As String = String.Empty
    Public Property windGustTimeUtcDailyStr As String = String.Empty
    Public Property windGustDirectionDaily As String = String.Empty ' Integer
    Public Property observationTimeAdjustedLocalStr As String = String.Empty
    Public Property feelsLike As String = String.Empty ' Double
  End Class

#End Region

#Region "Weatherbug Forecast"

  <Serializable()> _
  Public Class WeatherForecast
    Public Property dailyForecastPeriods As List(Of dailyForecastPeriod)
    Public Property forecastCreatedUtcStr As String = String.Empty
    Public Property location As String = String.Empty
    Public Property locationType As String = String.Empty
  End Class

  <Serializable()> _
  Public Class dailyForecastPeriod
    Public Property cloudCoverPercent As String = String.Empty ' Double
    Public Property dewPoint As String = String.Empty ' Double
    Public Property iconCode As String = String.Empty ' Integer
    Public Property precipCode As String = String.Empty ' Integer
    Public Property precipProbability As String = String.Empty ' Integer
    Public Property relativeHumidity As String = String.Empty ' Integer
    Public Property summaryDescription As String = String.Empty
    Public Property temperature As String = String.Empty ' Double
    Public Property thunderstormProbability As String = String.Empty ' Integer
    Public Property windDirectionDegrees As String = String.Empty ' Integer
    Public Property windSpeed As String = String.Empty ' Double
    Public Property detailedDescription As String = String.Empty
    Public Property forecastDateLocalStr As String = String.Empty
    Public Property forecastDateUtcStr As String = String.Empty
    Public Property isNightTimePeriod As String = String.Empty
  End Class

#End Region

#Region "Weather Alerts"

  <Serializable()> _
  Public Class WeatherAlerts
    Public Property alertList As List(Of WeatherAlert)
    Public Property location As String = String.Empty
    Public Property locationType As String = String.Empty
  End Class

  <Serializable()> _
  Public Class WeatherAlert
    Public Property AlertId As String = String.Empty
    Public Property AlertPrimaryId As String = String.Empty
    Public Property AlertProviderId As String = String.Empty
    Public Property AlertSecondaryId As String = String.Empty
    Public Property AlertType As String = String.Empty
    Public Property AlertTypeName As String = String.Empty
    Public Property ExpiredDateTimeLocalString As String = String.Empty
    Public Property ExpiredDateTimeUtcString As String = String.Empty
    Public Property IssuedDateTimeLocalString As String = String.Empty
    Public Property IssuedDateTimeUtcString As String = String.Empty
    Public Property Message As String = String.Empty
    Public Property Polygon As String = String.Empty
    Public Property RawText As String = String.Empty
    Public Property PVtec As String = String.Empty
  End Class

#End Region

#Region "Weatherbug Icons"

  <Serializable()> _
  Public Class WeatherIcons
    Public Property Width As String = String.Empty  ' Integer
    Public Property Height As String = String.Empty ' Integer
    Public Property Icons As List(Of WeatherIcon)
  End Class

  <Serializable()> _
  Public Class WeatherIcon
    Public Property URL As String = String.Empty
    Public Property IconCode As String = String.Empty ' Integer
  End Class

#End Region

End Class

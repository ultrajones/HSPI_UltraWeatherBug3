Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Net.Sockets
Imports System.Text
Imports System.Net
Imports System.Xml
Imports System.Data.Common
Imports System.Drawing
Imports HomeSeerAPI
Imports Scheduler
Imports System.ComponentModel

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable
  Const Pagename = "Events"

  Public HSDevices As New SortedList
  Public Stations As New SortedList
  Public Forecasts As New SortedList
  Public Alerts As New SortedList

  Private LastSpeak As New Hashtable()

  Public WeatherBugAPI As hspi_weatherbug_api
  Public WeatherBugAPILock As New Object
  Public gAPIKeyPrimary As String = String.Empty
  Public gAPIKeySecondary As String = String.Empty
  Public gAPICultureInfo As String = "en-us"

  Public Const IFACE_NAME As String = "UltraWeatherBug3"

  Public Const LINK_TARGET As String = "hspi_ultraweatherbug3/hspi_ultraweatherbug3.aspx"
  Public Const LINK_URL As String = "hspi_ultraweatherbug3.html"
  Public Const LINK_TEXT As String = "UltraWeatherBug3"
  Public Const LINK_PAGE_TITLE As String = "UltraWeatherBug3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultraweatherbug3/UltraWeatherBug3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = String.Empty
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultraweatherbug3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public gRTCReceived As Boolean = False
  Public gDeviceValueType As String = "1"
  Public gDeviceImage As Boolean = True
  Public gStatusImageSizeWidth As Integer = 32
  Public gStatusImageSizeHeight As Integer = 32

  Public gMonitoring As Boolean = True

  Public Const EMAIL_SUBJECT As String = "UltraWeatherBug3 $notification-type For $station"
  Public Const WEATHER_CURRENT_TEMPLATE As String = "Current weather conditions for $station:~" & _
                                                     "Temperature is $temp and $current-condition.~" & _
                                                     "Humidity is $humidity (feels like $feels-like).~" & _
                                                     "Winds $wind-direction at $wind-speed with~" & _
                                                     "wind gusts $gust-direction at $gust-speed."

  Public Const WEATHER_FORECAST_TEMPLATE As String = "Weather forecast for $station:~" & vbCrLf & _
                                                      "$forecast"

  Public Const WEATHER_ALERT_TEMPLATE As String = "Weather alert for $station:~" & _
                                                   "Posted on: $posted-date~" & _
                                                   "Expires on: $expires-date~" & _
                                                   "Type: $type~" & _
                                                   "Title: $title~" & _
                                                   "$msg-summary"

  Public Const SPEAK_WEATHER_CURRENT As String = "Current weather conditions for $station:~" & _
                                                  "Temperature is $temp and $current-condition.~" & _
                                                  "Humidity is $humidity (feels like $feels-like).~" & _
                                                  "Winds $wind-direction at $wind-speed with~" & _
                                                  "wind gusts $gust-direction at $gust-speed."

  Public Const SPEAK_WEATHER_FORECAST As String = "Weather forecast for $station:~" & _
                                                   "$forecast"

  Public Const SPEAK_WEATHER_ALERT As String = "Attention everyone, a $type has been issued for $stationcity.~" & _
                                                "$title."

#Region "UltraWeatherBug3 Public Functions"

  ''' <summary>
  ''' Returns various statistics
  ''' </summary>
  ''' <param name="StatisticsType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStatistics(ByVal StatisticsType As String) As Integer

    Select Case StatisticsType
      Case "StationCount"
        Return hspi_plugin.GetStationCount
      Case "AlertsCount"
        Return hspi_plugin.Alerts.Count
      Case "APISuccess"
        Return WeatherBugAPI.QuerySuccessCount()
      Case "APIFailure"
        Return WeatherBugAPI.QueryFailureCount()
      Case Else
        Return 0
    End Select

  End Function

  ''' <summary>
  ''' Returns stations sorted list
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStations() As SortedList

    Try
      If Stations.Count > 0 Then
        For index As Integer = 1 To 5
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          If Stations.ContainsKey(strStationNumber) = True Then
            ' Make sure we sync our device address with the current HomeSeer state
            For Each strKey As String In Stations(strStationNumber).Keys
              Stations(strStationNumber)(strKey)("DevCode") = GetDeviceAddress(String.Format("{0}-{1}", strStationNumber, strKey))
            Next
          End If
        Next
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetStations()")
    End Try

    Return hspi_plugin.Stations
  End Function

  ''' <summary>
  ''' Returns station sorted list
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStation(ByVal strStationNumber As String) As Hashtable

    Dim Station As New Hashtable

    Try
      If Stations.ContainsKey(strStationNumber) = True Then
        Dim strStationId As String = Trim(hs.GetINISetting(strStationNumber, "StationId", "", gINIFile))
        If strStationId.Length > 0 Then
          Station = Stations(strStationNumber).Clone
        End If
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetStation()")
    End Try

    Return Station

  End Function

  ''' <summary>
  ''' Returns forecasts sorted list
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetForecast(ByVal strStationNumber As String) As Hashtable

    Dim Forecast As New Hashtable

    Try
      If Forecasts.ContainsKey(strStationNumber) = True Then
        Dim strStationId As String = Trim(hs.GetINISetting(strStationNumber, "StationId", "", gINIFile))
        If strStationId.Length > 0 Then
          Forecast = Forecasts(strStationNumber).Clone
        End If
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetForecast()")
    End Try
    Return Forecast

  End Function

  ''' <summary>
  ''' Live Weather Thread
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub CheckWeatherThread()

    Dim strMessage As String = ""
    Dim iCheckInterval As Integer = 0

    Dim bAbortThread As Boolean = False

    Try
      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        If gMonitoring = True Then

          For index As Integer = 1 To 5
            Dim strStationNumber As String = String.Format("Station{0}", index.ToString)

            '
            ' Get the weather for the selected station
            '
            Thread.Sleep(1000)
            SyncLock WeatherBugAPILock
              GetLiveWeather(strStationNumber)
            End SyncLock

          Next

        End If

        '
        ' Sleep the requested number of minutes between runs
        '
        iCheckInterval = CInt(hs.GetINISetting("Options", "LiveWeatherUpdate", "5", gINIFile))
        Thread.Sleep(1000 * (60 * iCheckInterval))

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckWeatherThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckWeatherThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets weather data from WeatherBug stations
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <remarks></remarks>
  Public Sub GetLiveWeather(ByVal strStationNumber As String)

    Try

      If Stations.ContainsKey(strStationNumber) = False Then
        Exit Sub
      End If

      Dim strStationId As String = Trim(hs.GetINISetting(strStationNumber, "StationId", "", gINIFile))
      If strStationId.Length > 0 Then

        Dim RealTimeWeather As hspi_weatherbug_api.RealTimeWeather = WeatherBugAPI.GetRealtimeWeather(strStationId)
        If RealTimeWeather Is Nothing Then Exit Sub

        SetDeviceValue(strStationNumber, "ob-date", RealTimeWeather.observationTimeLocalStr)

        If GetIntegerValue(RealTimeWeather.iconCode) > 0 Then
          SetDeviceValue(strStationNumber, "current-condition", RealTimeWeather.iconCode)
        End If

        ' Removed temp-high, temp-low, wet-bulb
        SetDeviceValue(strStationNumber, "temp", RealTimeWeather.temperature)

        SetDeviceValue(strStationNumber, "temp-rate", RealTimeWeather.temperatureRate)
        SetDeviceValue(strStationNumber, "heat-index", RealTimeWeather.heatIndex)
        SetDeviceValue(strStationNumber, "feels-like", RealTimeWeather.feelsLike)
        SetDeviceValue(strStationNumber, "dew-point", RealTimeWeather.dewPoint)
        SetDeviceValue(strStationNumber, "dew-point-rate", RealTimeWeather.dewPointRate)

        ' Removed humidity-high, humidity-low
        SetDeviceValue(strStationNumber, "humidity", RealTimeWeather.humidity)
        SetDeviceValue(strStationNumber, "humidity-rate", RealTimeWeather.humidityRate)

        ' Removed pressure-high, pressure-low
        SetDeviceValue(strStationNumber, "pressure", RealTimeWeather.pressureSeaLevel)
        SetDeviceValue(strStationNumber, "pressure-rate", RealTimeWeather.pressureSeaLevelRate)

        '
        ' Removed rain-rate-max
        '
        SetDeviceValue(strStationNumber, "rain-month", RealTimeWeather.rainMonthly)
        SetDeviceValue(strStationNumber, "rain-rate", RealTimeWeather.rainRate)
        SetDeviceValue(strStationNumber, "rain-today", RealTimeWeather.rainDaily)
        SetDeviceValue(strStationNumber, "rain-year", RealTimeWeather.rainYearly)

        '
        ' Wind
        '
        SetDeviceValue(strStationNumber, "wind-direction", RealTimeWeather.windDirection)
        SetDeviceValue(strStationNumber, "wind-direction-avg", RealTimeWeather.windDirectionAvg)
        SetDeviceValue(strStationNumber, "wind-speed", RealTimeWeather.windSpeed)
        SetDeviceValue(strStationNumber, "wind-speed-avg", RealTimeWeather.windSpeedAvg)
        SetDeviceValue(strStationNumber, "gust-direction", RealTimeWeather.windGustDirectionHourly)
        SetDeviceValue(strStationNumber, "gust-speed", RealTimeWeather.windGustHourly)
        SetDeviceValue(strStationNumber, "gust-time", RealTimeWeather.windGustTimeLocalHourlyStr)

        '
        ' Removed light, light-rate
        '

        '
        ' Visibility
        '
        SetDeviceValue(strStationNumber, "visibility", RealTimeWeather.visibility)
        SetDeviceValue(strStationNumber, "visibility-rate", RealTimeWeather.visibilityRate)

      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      'ProcessError(pEx, "GetLiveWeather")
    End Try

  End Sub

  Private Function GetIntegerValue(ByVal Value As String) As Integer

    Dim RetVal As Integer = 0

    Try
      RetVal = Val(Value)
    Catch pEx As Exception

    End Try

    Return RetVal

  End Function

  Private Function GetDoubleValue(ByVal Value As String) As Double

    Dim RetVal As Double = 0

    Try
      RetVal = Val(Value)
    Catch pEx As Exception

    End Try

    Return RetVal

  End Function

  ''' <summary>
  ''' Weather Forecast Thread
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub CheckForecastThread()

    Dim strMessage As String = ""
    Dim iCheckInterval As Integer = 0

    Dim bAbortThread As Boolean = False

    Try
      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        If gMonitoring = True Then

          For index As Integer = 1 To 5
            Dim strStationNumber As String = String.Format("Station{0}", index.ToString)

            '
            ' Get the weather forecast for the selected station
            '
            Thread.Sleep(5000)
            SyncLock WeatherBugAPILock
              GetForecastWeather(strStationNumber)
            End SyncLock

          Next

        End If

        '
        ' Sleep the requested number of minutes between runs
        '
        iCheckInterval = CInt(hs.GetINISetting("Options", "ForecastUpdate", "30", gINIFile))
        Thread.Sleep(1000 * (60 * iCheckInterval))

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckForecastThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckForecastThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets weather data from WeatherBug stations
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <remarks></remarks>
  Public Sub GetForecastWeather(ByVal strStationNumber As String)

    Try

      If Forecasts.ContainsKey(strStationNumber) = False Then
        Exit Sub
      End If

      Dim strStationId As String = Trim(hs.GetINISetting(strStationNumber, "StationId", "", gINIFile))
      If strStationId.Length > 0 Then
        Dim i As Integer = 1

        Dim Today As String = DateTime.Now.DayOfWeek.ToString
        Dim Tomorrow As String = Now.AddDays(1).DayOfWeek.ToString

        Dim WeatherForecast As hspi_weatherbug_api.WeatherForecast = WeatherBugAPI.GetWeatherForecast(strStationId)
        If WeatherForecast Is Nothing Then Exit Sub

        Dim forecastCreatedDate As String = WeatherForecast.forecastCreatedUtcStr
        Dim location As String = WeatherForecast.location
        Dim locationType As String = WeatherForecast.locationType

        For Each dailyForecastPeriod As hspi_weatherbug_api.dailyForecastPeriod In WeatherForecast.dailyForecastPeriods

          Dim forecastDate As DateTime
          Date.TryParse(dailyForecastPeriod.forecastDateLocalStr, forecastDate)

          Dim dayDiff As Integer = DateDiff(DateInterval.Day, DateTime.Today, forecastDate)
          Dim isNightTimePeriod As Boolean = Boolean.Parse(dailyForecastPeriod.isNightTimePeriod)
          Dim ForecastDayNumber As Double = (dayDiff + 1) + IIf(isNightTimePeriod, 0.5, 0.0)
          Dim forecastDayName As String = ForecastDayNumber.ToString

          Dim dayOfWeek As String = forecastDate.DayOfWeek.ToString
          If dayOfWeek = Today Then
            dayOfWeek = IIf(isNightTimePeriod = True, "Tonight", "Today")
          ElseIf dayOfWeek = Tomorrow Then
            dayOfWeek = IIf(isNightTimePeriod = True, "Tomorrow Night", "Tomorrow")
          Else
            dayOfWeek &= IIf(isNightTimePeriod = True, " Night", "")
          End If

          If Forecasts(strStationNumber).ContainsKey(i) = True Then
            Dim Forecast As Hashtable = Forecasts(strStationNumber)(i)

            Forecast("title") = dayOfWeek
            Forecast("cloudCoverPercent") = String.Format("{0}%", dailyForecastPeriod.cloudCoverPercent)
            Forecast("dewPoint") = String.Format("{0} {1}", dailyForecastPeriod.dewPoint, GetDeviceSuffix("dewPoint", "Temperature"))
            Forecast("iconCode") = dailyForecastPeriod.iconCode
            Forecast("iconImage") = FormatDeviceImage("current-condition", dailyForecastPeriod.iconCode, 32, 32)
            Forecast("precipCode") = GetPrecipitationType(dailyForecastPeriod.precipCode)
            Forecast("precipProbability") = String.Format("{0}%", dailyForecastPeriod.precipProbability)
            Forecast("relativeHumidity") = String.Format("{0} {1}", dailyForecastPeriod.relativeHumidity, GetDeviceSuffix("relativeHumidity", "Humidity"))
            Forecast("summaryDescription") = dailyForecastPeriod.summaryDescription
            Forecast("detailedDescription") = dailyForecastPeriod.detailedDescription
            Forecast("temperature") = String.Format("{0} {1}", GetIntegerValue(dailyForecastPeriod.temperature), GetDeviceSuffix("dewPoint", "Temperature"))
            Forecast("thunderstormProbability") = String.Format("{0}%", dailyForecastPeriod.thunderstormProbability)
            Forecast("windDirectionDegrees") = GetWindDirectionShortName(dailyForecastPeriod.windDirectionDegrees)
            Forecast("windSpeed") = String.Format("{0} {1}", GetIntegerValue(dailyForecastPeriod.windSpeed), GetDeviceSuffix("wind-speed", "Wind"))
            Forecast("forecastDateLocalStr") = dailyForecastPeriod.forecastDateLocalStr

          End If

          Select Case dayDiff
            Case 0
              '
              ' This is Today
              ' 
              If isNightTimePeriod = False Then
                If i = 1 Then
                  SetDeviceValue(strStationNumber, "current-condition", dailyForecastPeriod.iconCode)
                End If

                SetDeviceValue(strStationNumber, "todays-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "todays-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                If i = 1 Then
                  SetDeviceValue(strStationNumber, "current-condition", dailyForecastPeriod.iconCode)
                End If

                SetDeviceValue(strStationNumber, "todays-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "todays-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
            Case 1
              '
              ' This is Tomorrow
              '
              If isNightTimePeriod = False Then
                SetDeviceValue(strStationNumber, "tomorrows-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "tomorrows-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                SetDeviceValue(strStationNumber, "tomorrows-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "tomorrows-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
            Case 2
              '
              ' 2-Day Forecast
              '
              If isNightTimePeriod = False Then
                SetDeviceValue(strStationNumber, "2-day-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "2-day-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                SetDeviceValue(strStationNumber, "2-day-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "2-day-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
            Case 3
              '
              ' 3-Day Forecast
              '
              If isNightTimePeriod = False Then
                SetDeviceValue(strStationNumber, "3-day-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "3-day-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                SetDeviceValue(strStationNumber, "3-day-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "3-day-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
            Case 4
              '
              ' 3-Day Forecast
              '
              If isNightTimePeriod = False Then
                SetDeviceValue(strStationNumber, "4-day-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "4-day-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                SetDeviceValue(strStationNumber, "4-day-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "4-day-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
            Case 5
              '
              ' 3-Day Forecast
              '
              If isNightTimePeriod = False Then
                SetDeviceValue(strStationNumber, "5-day-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "5-day-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                SetDeviceValue(strStationNumber, "5-day-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "5-day-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
            Case 6
              '
              ' 3-Day Forecast
              '
              If isNightTimePeriod = False Then
                SetDeviceValue(strStationNumber, "6-day-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "6-day-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                SetDeviceValue(strStationNumber, "6-day-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "6-day-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
            Case 7
              '
              ' 3-Day Forecast
              '
              If isNightTimePeriod = False Then
                SetDeviceValue(strStationNumber, "7-day-temperature-day", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "7-day-short-prediction-day", dailyForecastPeriod.iconCode)
              Else
                SetDeviceValue(strStationNumber, "7-day-temperature-night", dailyForecastPeriod.temperature)
                SetDeviceValue(strStationNumber, "7-day-short-prediction-night", dailyForecastPeriod.iconCode)
              End If
          End Select

          '
          ' Increment the Forecast Index
          '
          i += 1

        Next

      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      'ProcessError(pEx, "GetForecastWeather")
    End Try

  End Sub

  ''' <summary>
  ''' Weather Alerts Thread
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub CheckAlertsThread()

    Dim strMessage As String = ""
    Dim iCheckInterval As Integer = 0

    Dim bAbortThread As Boolean = False

    Try
      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        '
        ' Clear expired alerts
        '
        ClearExpiredWeatherAlerts()

        If gMonitoring = True Then

          For index As Integer = 1 To 5
            Dim strStationNumber As String = String.Format("Station{0}", index.ToString)

            '
            ' Get the weather forecast for the selected station
            '
            Thread.Sleep(10000)
            SyncLock WeatherBugAPILock
              GetAlertsWeather(strStationNumber)
            End SyncLock

          Next

        End If

        '
        ' Update the Alerts count
        '
        Dim dv_addr As String = String.Concat(IFACE_NAME, "-Alerts")
        Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
        Dim bDeviceExists As Boolean = dv_ref <> -1

        If bDeviceExists = True Then
          '
          ' Update the HomeSeer device
          '
          Dim iDeviceValue As Integer = Alerts.Count

          If hs.DeviceValue(dv_ref) <> iDeviceValue Then
            hs.SetDeviceValueByRef(dv_ref, iDeviceValue, True)
          End If

        End If

        '
        ' Sleep the requested number of minutes between runs
        '
        iCheckInterval = CInt(hs.GetINISetting("Options", "AlertsUpdate", "30", gINIFile))
        Thread.Sleep(1000 * (60 * iCheckInterval))

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckAlertsThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckAlertsThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets weather data from WeatherBug stations
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <remarks></remarks>
  Public Sub GetAlertsWeather(ByVal strStationNumber As String)

    Try

      Dim unitType As String = hs.GetINISetting("Options", "UnitType", "0", gINIFile)
      Dim bAlertsOffset As Boolean = CBool(hs.GetINISetting("Options", "AlertsOffset", "True", gINIFile))

      Dim strStationId As String = Trim(hs.GetINISetting(strStationNumber, "StationId", "", gINIFile))
      If strStationId.Length > 0 Then

        Dim WeatherAlerts As hspi_weatherbug_api.WeatherAlerts = WeatherBugAPI.GetWeatherAlerts(strStationId)
        If WeatherAlerts Is Nothing Then Exit Sub

        For Each WeatherAlert As hspi_weatherbug_api.WeatherAlert In WeatherAlerts.alertList
          Dim strAlertId As String = WeatherAlert.AlertId

          Dim strPostedDate As String = WeatherAlert.IssuedDateTimeLocalString
          Dim strExpiresDate As String = WeatherAlert.ExpiredDateTimeLocalString

          Dim strAlertType As String = WeatherBugAPI.GetAlertSignificance(WeatherAlert.AlertType)
          Dim strAlertTitle As String = WeatherAlert.AlertTypeName
          Dim strAlertSummary As String = WeatherAlert.Message.Replace("\r\n", vbCrLf)

          SyncLock Alerts.SyncRoot

            If Alerts.ContainsKey(strAlertId) = False Then

              Alerts.Add(strAlertId, New Hashtable())

              For index As Integer = 1 To 6
                Dim strStationNum As String = String.Format("Station{0}", index.ToString)
                Alerts(strAlertId)(strStationNum) = False
              Next

              '
              ' Update the last alert type and title
              '
              SetDeviceValue(strStationNumber, "last-alert-type", strAlertType)
              SetDeviceValue(strStationNumber, "last-alert-title", strAlertTitle)

              Try

                Alerts(strAlertId)("expires-date") = strExpiresDate
                Alerts(strAlertId)("posted-date") = strPostedDate
                Alerts(strAlertId)("type") = strAlertType
                Alerts(strAlertId)("title") = strAlertTitle
                Alerts(strAlertId)("msg-summary") = strAlertSummary

                '
                ' See if we need to trigger an alert
                '
                If Alerts(strAlertId)(strStationNumber) = False Then
                  Alerts(strAlertId)(strStationNumber) = True

                  Dim strTriggerAny As String = String.Format("{0},{1},{2}", "Weather Alert Trigger", strStationNumber, "Any")
                  CheckTrigger(IFACE_NAME, WeatherTriggers.WeatherAlert, -1, strTriggerAny)

                  '
                  ' Loop through the WeatherAlerts members
                  '
                  Dim names As String() = System.Enum.GetNames(GetType(WeatherAlertTypes))
                  For i As Integer = 0 To names.Length - 1
                    If Regex.IsMatch(strAlertType, names(i), RegexOptions.IgnoreCase) = True Then
                      '
                      ' Check weather trigger
                      '
                      Dim strTrigger As String = String.Format("{0},{1},{2}", "Weather Alert Trigger", strStationNumber, names(i))
                      CheckTrigger(IFACE_NAME, WeatherTriggers.WeatherAlert, -1, strTrigger)
                    End If
                  Next

                End If

              Catch pEx As Exception
                '
                ' Process the error
                '
                WriteMessage("GetAlertsWeather failed due to an error: " & pEx.Message, MessageType.Debug)
              End Try

            End If

          End SyncLock

        Next

      Else
        '
        ' No weather alerts
        '
      End If

      '
      ' Clear the last alert device
      '
      If Alerts.Count = 0 Then
        SetDeviceValue(strStationNumber, "last-alert-type", "N/A")
        SetDeviceValue(strStationNumber, "last-alert-title", "N/A")
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      'ProcessError(pEx, "GetAlertsWeather")
    End Try

  End Sub

  ''' <summary>
  ''' Clears expired weather alerts
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ClearExpiredWeatherAlerts()

    Try

      Dim purgeQueue As New Queue

      SyncLock Alerts.SyncRoot

        For Each strAlertId As String In Alerts.Keys

          Dim strExpiresDate As String = Alerts(strAlertId)("expires-date")
          Dim expiresDate As DateTime
          If DateTime.TryParse(strExpiresDate, expiresDate) = True Then
            Dim lngDiff As Long = DateDiff(DateInterval.Hour, expiresDate, Date.Now)
            WriteMessage("Date difference is " & lngDiff.ToString, MessageType.Debug)
            If DateDiff(DateInterval.Second, expiresDate, Date.Now) > 0 Then
              purgeQueue.Enqueue(strAlertId)
            End If
          End If

        Next

        '
        ' Purge the expired alerts
        '
        If purgeQueue.Count > 0 Then
          While purgeQueue.Count > 0
            Dim strAlertId As String = purgeQueue.Dequeue
            If Alerts.ContainsKey(strAlertId) Then
              Alerts.Remove(strAlertId)
            End If
          End While
        End If

      End SyncLock

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ClearExpiredWeatherAlerts()")
    End Try

  End Sub

  ''' <summary>
  ''' E-Mails current weather conditions
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <remarks></remarks>
  Public Sub EmailWeatherConditions(ByVal strStationNumber As String)

    Try

      Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "", gINIFile))

      Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "smtp_from", "")
      Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "smtp_to", "")

      Dim strEmailTo As String = hs.GetINISetting("EmailNotification", "EmailRcptTo", strEmailRcptTo, gINIFile)
      Dim strEmailFrom As String = hs.GetINISetting("EmailNotification", "EmailFrom", strEmailFromDefault, gINIFile)
      Dim strEmailSubject As String = hs.GetINISetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT, gINIFile)
      Dim strEmailBody As String = hs.GetINISetting("EmailNotification", "WeatherCurrent", WEATHER_CURRENT_TEMPLATE, gINIFile)

      If Stations.ContainsKey(strStationNumber) = False Then
        Throw New Exception("Unable to send weather conditions for station because the station number was not found.")
      ElseIf Regex.IsMatch(strEmailFrom, ".+@.+") = False Then
        Throw New Exception("Unable to send weather conditions for station because the sender is not a valid e-mail address.")
      ElseIf Regex.IsMatch(strEmailTo, ".+@.+") = False Then
        Throw New Exception("Unable to send weather conditions for station because the recipient is not a valid e-mail address.")
      End If

      '
      ' Do custom variable replacment
      '
      strEmailSubject = strEmailSubject.Replace("$stationname", strStationName)
      strEmailSubject = strEmailSubject.Replace("$station", strStationNumber)
      strEmailSubject = strEmailSubject.Replace("$notification-type", "Current Conditions")

      strEmailBody = strEmailBody.Replace("$notification-type", "Current Conditions")
      strEmailBody = ReplaceVariables(strStationNumber, strEmailBody)
      strEmailBody = strEmailBody.Replace("&deg;", "°").Replace("~", vbCrLf)

      Dim List() As String = hs.GetPluginsList()
      If List.Contains("UltraSMTP3:") = True Then
        '
        ' Send e-mail using UltraSMTP3
        '
        hs.PluginFunction("UltraSMTP3", "", "SendMail", New Object() {strEmailTo, strEmailSubject, strEmailBody, Nothing})
      Else
        '
        ' Send e-mail using HomeSeer
        '
        hs.SendEmail(strEmailTo, strEmailFrom, "", "", strEmailSubject, strEmailBody, "")
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EmailWeatherConditions()")
    End Try

  End Sub

  ''' <summary>
  ''' E-Mails current weather forecast
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <remarks></remarks>
  Public Sub EmailWeatherForecast(ByVal strStationNumber As String)

    Try

      Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "", gINIFile))

      Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "smtp_from", "")
      Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "smtp_to", "")

      Dim strEmailTo As String = hs.GetINISetting("EmailNotification", "EmailRcptTo", strEmailRcptTo, gINIFile)
      Dim strEmailFrom As String = hs.GetINISetting("EmailNotification", "EmailFrom", strEmailFromDefault, gINIFile)
      Dim strEmailSubject As String = hs.GetINISetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT, gINIFile)
      Dim strEmailBody As String = hs.GetINISetting("EmailNotification", "WeatherForecast", WEATHER_FORECAST_TEMPLATE, gINIFile)

      If Forecasts.ContainsKey(strStationNumber) = False Then
        Throw New Exception("Unable to send weather forecast for station because the station number was not found.")
      ElseIf Regex.IsMatch(strEmailFrom, ".+@.+") = False Then
        Throw New Exception("Unable to send weather forecast for station because the sender is not a valid e-mail address.")
      ElseIf Regex.IsMatch(strEmailTo, ".+@.+") = False Then
        Throw New Exception("Unable to send weather forecast for station because the recipient is not a valid e-mail address.")
      End If

      '
      ' Build forecast for next n days
      '
      Dim iForecast As Integer = 1
      Dim iDays As Integer = 7
      Dim regexPattern As String = "\$forecast(?<days>\d)"
      Dim Matches As Match = Regex.Match(strEmailBody, regexPattern)
      If Matches.Length > 1 Then
        iDays = CInt(Matches.Groups("days").ToString)
      End If

      Dim weatherForecast As New StringBuilder
      For i As Integer = 1 To 14
        If iForecast <= iDays * 2 Then

          Dim Forecast As Hashtable = Forecasts(strStationNumber)(i)
          Dim strDayName As String = Forecast("title")

          weatherForecast.Append(String.Format("{0}, {1}{2}", strDayName, Forecast("detailedDescription"), vbCrLf))
        End If
        iForecast += 1
      Next

      '
      ' Do custom variable replacment
      '
      strEmailSubject = strEmailSubject.Replace("$stationname", strStationName)
      strEmailSubject = strEmailSubject.Replace("$station", strStationNumber)
      strEmailSubject = strEmailSubject.Replace("$notification-type", "Forecast")

      strEmailBody = strEmailBody.Replace("$notification-type", "Forecast")
      strEmailBody = ReplaceVariables(strStationNumber, strEmailBody)
      strEmailBody = strEmailBody.Replace("&deg;", "°").Replace("~", vbCrLf)

      strEmailBody = Regex.Replace(strEmailBody, "\$forecast(\d)?", weatherForecast.ToString)

      Dim List() As String = hs.GetPluginsList()
      If List.Contains("UltraSMTP3:") = True Then
        '
        ' Send e-mail using UltraSMTP3
        '
        hs.PluginFunction("UltraSMTP3", "", "SendMail", New Object() {strEmailTo, strEmailSubject, strEmailBody, Nothing})
      Else
        '
        ' Send e-mail using HomeSeer
        '
        hs.SendEmail(strEmailTo, strEmailFrom, "", "", strEmailSubject, strEmailBody, "")
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EmailWeatherForecast()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   E-Mails current weather alerts
  'Input:     Station Number as String
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub EmailWeatherAlerts(ByVal strStationNumber As String)

    Try

      If Stations.ContainsKey(strStationNumber) = False Then
        Throw New Exception("Unable to send weather alerts for station because the station number was not found.")
      End If

      Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "", gINIFile))

      Dim strEmailFromDefault As String = hs.GetINISetting("Settings", "smtp_from", "")
      Dim strEmailRcptTo As String = hs.GetINISetting("Settings", "smtp_to", "")
      Dim strEmailTo As String = hs.GetINISetting("EmailNotification", "EmailRcptTo", strEmailRcptTo, gINIFile)
      Dim strEmailFrom As String = hs.GetINISetting("EmailNotification", "EmailFrom", strEmailFromDefault, gINIFile)

      SyncLock Alerts.SyncRoot

        For Each strAlertID As String In Alerts.Keys
          Dim Alert As Hashtable = Alerts(strAlertID)
          If Alert(strStationNumber) = True Then

            Dim strEmailSubject As String = hs.GetINISetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT, gINIFile)
            Dim strEmailBody As String = hs.GetINISetting("EmailNotification", "WeatherAlert", WEATHER_ALERT_TEMPLATE, gINIFile)

            strEmailBody = strEmailBody.Replace("$posted-date", Alert("posted-date"))
            strEmailBody = strEmailBody.Replace("$expires-date", Alert("expires-date"))
            strEmailBody = strEmailBody.Replace("$type", Alert("type"))
            strEmailBody = strEmailBody.Replace("$title", Alert("title"))
            strEmailBody = strEmailBody.Replace("$msg-summary", Alert("msg-summary"))

            strEmailSubject = strEmailSubject.Replace("$stationname", strStationName)
            strEmailSubject = strEmailSubject.Replace("$station", strStationNumber)
            strEmailSubject = strEmailSubject.Replace("$notification-type", "Alerts")

            strEmailBody = strEmailBody.Replace("$notification-type", "Alerts")
            strEmailBody = ReplaceVariables(strStationNumber, strEmailBody)
            strEmailBody = strEmailBody.Replace("&deg;", "°").Replace("~", vbCrLf)

            Dim List() As String = hs.GetPluginsList()
            If List.Contains("UltraSMTP3:") = True Then
              '
              ' Send e-mail using UltraSMTP3
              '
              hs.PluginFunction("UltraSMTP3", "", "SendMail", New Object() {strEmailTo, strEmailSubject, strEmailBody, Nothing})
            Else
              '
              ' Send e-mail using HomeSeer
              '
              hs.SendEmail(strEmailTo, strEmailFrom, "", "", strEmailSubject, strEmailBody, "")
            End If
          End If

        Next

      End SyncLock

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EmailWeatherAlerts()")
    End Try

  End Sub

  ''' <summary>
  ''' Speak current weather conditions
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="bWait"></param>
  ''' <param name="strSpeakerHost"></param>
  ''' <remarks></remarks>
  Public Sub SpeakWeatherConditions(ByVal strStationNumber As String, Optional ByVal bWait As Boolean = False, Optional ByVal strSpeakerHost As String = "")

    Try

      If Stations.ContainsKey(strStationNumber) = False Then
        Throw New Exception("Unable to send weather conditions for station because the station number was not found.")
      End If

      Dim strOutput As String = hs.GetINISetting("Speak", "WeatherCurrent", SPEAK_WEATHER_CURRENT, gINIFile)

      '
      ' Do custom variable replacment
      '
      strOutput = ReplaceVariables(strStationNumber, strOutput)
      strOutput = strOutput.Replace("$notification-type", "Current Conditions")

      '
      ' Speak the weather
      '
      SpeakWeather(strOutput, bWait, strSpeakerHost)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "EmailWeatherConditions()")
    End Try

  End Sub

  ''' <summary>
  ''' Speaks the current weather forecast
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="bWait"></param>
  ''' <param name="strSpeakerHost"></param>
  ''' <remarks></remarks>
  Public Sub SpeakWeatherForecast(ByVal strStationNumber As String, Optional ByVal bWait As Boolean = False, Optional ByVal strSpeakerHost As String = "")

    Try

      If Forecasts.ContainsKey(strStationNumber) = False Then
        Throw New Exception("Unable to send weather forecast for station because the station number was not found.")
      End If

      Dim strOutput As String = hs.GetINISetting("Speak", "WeatherForecast", SPEAK_WEATHER_FORECAST, gINIFile)

      '
      ' Do custom variable replacment
      '
      strOutput = WeatherForecastReplacmentVariables(strStationNumber, strOutput)

      '
      ' Speak the Weather Forecast
      '
      SpeakWeather(strOutput, bWait, strSpeakerHost)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SpeakWeatherForecast()")
    End Try

  End Sub

  ''' <summary>
  ''' Speak current weather alerts
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="bWait"></param>
  ''' <param name="strSpeakerHost"></param>
  ''' <remarks></remarks>
  Public Sub SpeakWeatherAlerts(ByVal strStationNumber As String, Optional ByVal bWait As Boolean = False, Optional ByVal strSpeakerHost As String = "")

    Try

      If Forecasts.ContainsKey(strStationNumber) = False Then
        Throw New Exception("Unable to send weather alerts for station because the station number was not found.")
      End If

      SyncLock Alerts.SyncRoot

        For Each strAlertID As String In Alerts.Keys

          Dim Alert As Hashtable = Alerts(strAlertID)
          If Alert(strStationNumber) = True Then

            Dim strOutput As String = hs.GetINISetting("Speak", "WeatherAlert", SPEAK_WEATHER_ALERT, gINIFile)

            strOutput = WeatherAlertsReplacementVariables(strStationNumber, strOutput)

            '
            ' Do custom variable replacment
            '
            strOutput = ReplaceVariables(strStationNumber, strOutput)
            strOutput = strOutput.Replace("$notification-type", "Alerts")

            '
            ' Speak the weather alert
            '
            SpeakWeather(strOutput, bWait, strSpeakerHost)

          End If

        Next

      End SyncLock

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SpeakWeatherAlerts()")
    End Try

  End Sub

  ''' <summary>
  ''' Fixes up the speak output
  ''' </summary>
  ''' <param name="weather"></param>
  ''' <param name="bWait"></param>
  ''' <param name="strSpeakerHost"></param>
  ''' <remarks></remarks>
  Private Sub SpeakWeather(ByVal weather As String, Optional ByVal bWait As Boolean = False, Optional ByVal strSpeakerHost As String = "")

    Try
      '
      ' Ensure the speaker host is in the correct format
      '
      If Regex.IsMatch(strSpeakerHost, ".+:(\*|.+)") = False Then
        strSpeakerHost = "*:*"
      End If

      If strSpeakerHost = "Default:*" Then
        strSpeakerHost = "*:*"
      End If

      '
      ' Prevent Duplicate Output
      '
      Dim MyTimer As Long = DateAndTime.Timer
      Dim MD5Sum As String = GenerateHash(strSpeakerHost & weather)
      If LastSpeak.ContainsKey(MD5Sum) = False Then
        LastSpeak.Add(MD5Sum, MyTimer)
      Else
        Dim lastSceen As Long = LastSpeak(MD5Sum)
        Dim diff As Long = CLng(MyTimer - lastSceen)
        LastSpeak(MD5Sum) = MyTimer
        If diff <= 30 Then
          Return
        End If
      End If

      '
      ' Clean up the Weather Variables
      '
      weather = SpeachReplacementVariables(weather)

      '
      ' Speak the weather
      '
      hs.Speak(weather, bWait, strSpeakerHost)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SpeakWeather()")
    End Try

  End Sub

  ''' <summary>
  ''' Process Replacment Variables
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="strInputString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ReplaceVariables(ByVal strStationNumber As String, ByVal strInputString As String) As String

    Dim strOutput As String = strInputString

    Try

      Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "", gINIFile))
      Dim strCityName As String = Trim(hs.GetINISetting(strStationNumber, "CityName", "", gINIFile))
      strCityName = Regex.Replace(strCityName, ",.+", "")

      '
      ' Get the station hashtable
      '
      Dim Station As Hashtable = Stations(strStationNumber)

      '
      ' Do variable replacment
      '
      Dim KeyNames() As String = {"Weather", "Temperature", "Humidity", "Wind", "Rain", "Pressure", "Visibility", "Forecast", "Alerts"}

      For Each strKeyName As String In KeyNames
        '
        ' Get the list of xml nodes (keys) we are interested in
        '
        Dim Keys() As String = GetWeatherKeys(strKeyName)

        '
        ' Go through each pre-defined XML node element (key)
        '
        For Each strKey As String In Keys
          If Stations(strStationNumber).ContainsKey(strKey) = True Then
            Dim strPlaceholder As String = String.Format("${0}", strKey)
            strOutput = strOutput.Replace(strPlaceholder, FormatDeviceValue(strKey, Station(strKey)))
          End If
        Next
      Next

      strOutput = strOutput.Replace("$stationname", strStationName)
      strOutput = strOutput.Replace("$stationcity", strCityName)
      strOutput = strOutput.Replace("$station", strStationNumber)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ReplaceVariables()")
    End Try

    Return strOutput

  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="device"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function FormatDeviceValue(device_key As String, device As Hashtable) As String

    Try

      If device("Units").Length > 0 Then
        Return String.Format("{0}{1}", device("Value"), device("Units"))
      ElseIf device("String").Length > 0 Then
        Return device("String")
      Else
        Return "--"
      End If

    Catch pEx As Exception
      Return device("Value")
    End Try

  End Function

  ''' <summary>
  ''' Weather Forecast Replacment Variables
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="strInputString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WeatherForecastReplacmentVariables(ByVal strStationNumber As String, ByVal strInputString As String)

    Dim strOutput As String = strInputString

    Try
      '
      ' Build forecast for next n days
      '
      Dim iForecast As Integer = 1
      Dim iDays As Integer = 7
      Dim regexPattern As String = "\$forecast(?<days>\d)"
      Dim Matches As Match = Regex.Match(strOutput, regexPattern)
      If Matches.Length > 1 Then
        iDays = CInt(Matches.Groups("days").ToString)
      End If

      Dim weatherForecast As New StringBuilder
      For i As Integer = 1 To 14
        If iForecast <= iDays * 2 Then

          Dim Forecast As Hashtable = Forecasts(strStationNumber)(i)
          Dim strDayName As String = Forecast("title")

          weatherForecast.Append(String.Format("{0}, {1}{2}", strDayName, Forecast("detailedDescription"), vbCrLf))
        End If
        iForecast += 1
      Next

      strOutput = Regex.Replace(strOutput, "\$forecast(\d)?", weatherForecast.ToString)
      strOutput = ReplaceVariables(strStationNumber, strOutput)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "WeatherForecastReplacmentVariables()")
    End Try

    Return strOutput

  End Function

  ''' <summary>
  ''' Weather Alerts Replacement Variables
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="strInputString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WeatherAlertsReplacementVariables(ByVal strStationNumber As String, ByVal strInputString As String)

    Dim strOutput As String = strInputString

    Try

      SyncLock Alerts.SyncRoot

        For Each strAlertID As String In Alerts.Keys

          Dim Alert As Hashtable = Alerts(strAlertID)
          If Alert(strStationNumber) = True Then

            strOutput = strOutput.Replace("$posted-date", Alert("posted-date"))
            strOutput = strOutput.Replace("$expires-date", Alert("expires-date"))
            strOutput = strOutput.Replace("$type", Alert("type"))
            strOutput = strOutput.Replace("$title", Alert("title"))
            strOutput = strOutput.Replace("$msg-summary", Alert("msg-summary"))

          End If

        Next

      End SyncLock

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "WeatherAlertsReplacementVariables()")
    End Try

    Return strOutput

  End Function

  ''' <summary>
  ''' Speach Replacement Variables
  ''' </summary>
  ''' <param name="strInputString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SpeachReplacementVariables(ByVal strInputString As String)

    Dim strOutput As String = strInputString

    Try

      strOutput = Regex.Replace(strOutput, "~", " ")
      strOutput = Regex.Replace(strOutput, "&deg;(F|C)", " degrees ")
      strOutput = Regex.Replace(strOutput, "-- degrees", "unknown")
      strOutput = Regex.Replace(strOutput, "mph", "miles per hour")
      strOutput = Regex.Replace(strOutput, "\s+", " ")
      strOutput = Regex.Replace(strOutput, "\s+%", "%")
      strOutput = Regex.Replace(strOutput, "[.]{2,}", ", ")
      strOutput = Regex.Replace(strOutput, "\s[.]\s?", ". ")
      strOutput = Regex.Replace(strOutput, "\s[,]\s?", ",")
      strOutput = Regex.Replace(strOutput, "\s[;]\s?", ";")

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SpeachReplacementVariables()")
    End Try

    Return strOutput

  End Function

  ''' <summary>
  ''' Returns the station count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStationCount() As Integer

    Dim iStationCount As Integer = 0

    Try

      If Stations.Count > 0 Then
        For index As Integer = 1 To 5
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationId As String = Trim(hs.GetINISetting(strStationNumber, "StationId", "", gINIFile))
          If strStationId.Length > 0 Then
            iStationCount += 1
          End If
        Next
      End If
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetStationCount()")
    End Try

    Return iStationCount

  End Function

  ''' <summary>
  ''' Returns the days of the week
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDaysOfWeek() As ArrayList

    Dim DaysOfWeek As New ArrayList

    Try

      For value As Integer = 0 To 6

        Dim TheDate As DateTime = Now.AddDays(value)
        Dim DayOfWeek As New Hashtable

        Dim strDayOfWeek As String = TheDate.DayOfWeek.ToString
        Dim strDayOfWeekShort As String = strDayOfWeek.Substring(0, 3).ToUpper
        Dim strDayOfForecast As String = String.Format("{0} Day", value.ToString)

        If value = 0 Then strDayOfWeek = "Today"
        If value = 1 Then strDayOfWeek = "Tomorrow"

        DayOfWeek.Add("ForecastName", strDayOfForecast)
        DayOfWeek.Add("ShortName", strDayOfWeekShort)
        DayOfWeek.Add("LongName", strDayOfWeek)
        DaysOfWeek.Add(DayOfWeek)

      Next

      Return DaysOfWeek
    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "GetDaysOfWeek()")
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Returns the wind value
  ''' </summary>
  ''' <param name="strWindDirection"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWindDirectionValue(ByVal strWindDirection As String) As Double

    Try

      Dim windDirection As Integer = Val(strWindDirection)

      Select Case windDirection
        Case 0 To 11 : Return 0
        Case 12 To 34 : Return 22.5
        Case 35 To 56 : Return 45
        Case 57 To 78 : Return 67.5
        Case 79 To 101 : Return 90
        Case 102 To 123 : Return 112.5
        Case 123 To 146 : Return 135
        Case 147 To 168 : Return 157.5
        Case 169 To 191 : Return 180
        Case 192 To 213 : Return 202.5
        Case 214 To 236 : Return 225
        Case 237 To 258 : Return 247.5
        Case 259 To 291 : Return 270
        Case 282 To 303 : Return 292.5
        Case 204 To 236 : Return 315
        Case 237 To 348 : Return 337.5
        Case 349 To 359 : Return 0
        Case Else : Return -1
      End Select

    Catch pEx As Exception
      Return -1
    End Try

  End Function

  ''' <summary>
  ''' Returns the wind value
  ''' </summary>
  ''' <param name="strWindDirection"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWindDirectionShortName(ByVal strWindDirection As String) As String

    Try

      Dim windDirection As Integer = Val(strWindDirection)

      Select Case windDirection
        Case 0 To 11 : Return "N"
        Case 12 To 34 : Return "NNE"
        Case 35 To 56 : Return "NE"
        Case 57 To 78 : Return "ENE"
        Case 79 To 101 : Return "E"
        Case 102 To 123 : Return "ESE"
        Case 123 To 146 : Return "SE"
        Case 147 To 168 : Return "SSE"
        Case 169 To 191 : Return "S"
        Case 192 To 213 : Return "SSW"
        Case 214 To 236 : Return "SW"
        Case 237 To 258 : Return "WSW"
        Case 259 To 291 : Return "W"
        Case 282 To 303 : Return "WNW"
        Case 304 To 326 : Return "NW"
        Case 327 To 348 : Return "NNW"
        Case 349 To 359 : Return "N"
        Case Else : Return -1
      End Select

    Catch ex As Exception
      Return -1
    End Try

  End Function

  ''' <summary>
  ''' Returns the wind value
  ''' </summary>
  ''' <param name="strWindDirection"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWindDirectionLongName(ByVal strWindDirection As String) As String

    Try

      Dim windDirection As Integer = Val(strWindDirection)

      Select Case windDirection
        Case 0 To 11 : Return "North"
        Case 12 To 34 : Return "North Northeast"
        Case 35 To 56 : Return "NorthEast"
        Case 57 To 78 : Return "East Northeast"
        Case 79 To 101 : Return "East"
        Case 102 To 123 : Return "East Southeast"
        Case 123 To 146 : Return "Southeast"
        Case 147 To 168 : Return "South Southeast"
        Case 169 To 191 : Return "South"
        Case 192 To 213 : Return "South Southwest"
        Case 214 To 236 : Return "Southwest"
        Case 237 To 258 : Return "West Southwest"
        Case 259 To 291 : Return "West"
        Case 282 To 303 : Return "West Northwest"
        Case 304 To 326 : Return "Northwest"
        Case 327 To 348 : Return "North Northwest"
        Case 349 To 359 : Return "North"
        Case Else : Return -1
      End Select

    Catch ex As Exception
      Return strWindDirection
    End Try

  End Function

  ''' <summary>
  ''' Returns the wind value
  ''' </summary>
  ''' <param name="strWindDirection"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWindValue(ByVal strWindDirection As String) As Double

    Try

      Select Case strWindDirection
        Case "N", "North" : Return 0
        Case "NNE", "NorthNorthEast" : Return 22.5
        Case "NE", "NorthEast" : Return 45
        Case "ENE", "EastNorthEast" : Return 67.5
        Case "E", "East" : Return 90
        Case "ESE", "EastSouthEast" : Return 112.5
        Case "SE", "SouthEast" : Return 135
        Case "SSE", "SouthSouthEast" : Return 157.5
        Case "S", "South" : Return 180
        Case "SSW", "SouthSouthwest" : Return 202.5
        Case "SW", "Southwest" : Return 225
        Case "WSW", "WestSouthwest" : Return 247.5
        Case "W", "West" : Return 270
        Case "WNW", "WestNorthwest" : Return 292.5
        Case "NW", "Northwest" : Return 315
        Case "NNW", "NorthNorthwest" : Return 337.5
        Case Else : Return -1
      End Select

    Catch ex As Exception
      Return -1
    End Try

  End Function

  ''' <summary>
  ''' Returns the wind value
  ''' </summary>
  ''' <param name="strWindDirection"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWindDirection(ByVal strWindDirection As String) As String

    Try

      Select Case strWindDirection.ToUpper
        Case "N" : Return "North"
        Case "NNE" : Return "North Northeast"
        Case "NE" : Return "NorthEast"
        Case "ENE" : Return "East Northeast"
        Case "E" : Return "East"
        Case "ESE" : Return "East Southeast"
        Case "SE" : Return "Southeast"
        Case "SSE" : Return "South Southeast"
        Case "S" : Return "South"
        Case "SSW" : Return "South Southwest"
        Case "SW" : Return "Southwest"
        Case "WSW" : Return "West Southwest"
        Case "W" : Return "West"
        Case "WNW" : Return "West Northwest"
        Case "NW" : Return "Northwest"
        Case "NNW" : Return "NorthNorthwest"
        Case Else : Return strWindDirection
      End Select

    Catch ex As Exception
      Return strWindDirection
    End Try

  End Function

  Public Function GetPrecipitationType(precipCode As Integer)

    Select Case precipCode
      Case 0 : Return "No Precipitation"
      Case 1 : Return "Rain"
      Case 2 : Return "Snow"
      Case 3 : Return "Rain-snow Mix"
      Case 4 : Return "Sleet"
      Case 5 : Return "Freezing Rain"
      Case 6 : Return "Frozen Mix"
    End Select

  End Function

  ''' <summary>
  ''' Returns the Weather Condtion Name by Id
  ''' </summary>
  ''' <param name="iCondition"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWeatherConditionNameById(iCondition As Integer)

    Select Case iCondition
      Case 0 : Return "Clear"
      Case 1 : Return "Cloudy"
      Case 2 : Return "Partly Cloudy"
      Case 3 : Return "Partly Cloudy"
      Case 4 : Return "Partly Sunny"
      Case 5 : Return "Rain"
      Case 6 : Return "Thunderstorms"
      Case 7 : Return "Sunny"
      Case 8 : Return "Snow"
      Case 9 : Return "Flurries"
      Case 10 : Return "Unknown"
      Case 11 : Return "Chance of Snow"
      Case 12 : Return "Snow"
      Case 13 : Return "Cloudy"
      Case 14 : Return "Rain"
      Case 15 : Return "Chance of Rain"
      Case 16 : Return "Partly Cloudy"
      Case 17 : Return "Clear"
      Case 18 : Return "Thunderstorms"
      Case 19 : Return "Chance of Flurries"
      Case 20 : Return "Chance of Rain"
      Case 21 : Return "Chance of Sleet"
      Case 22 : Return "Chance of Storms"
      Case 23 : Return "Hazy"
      Case 24 : Return "Mostly Cloudy"
      Case 25 : Return "Sleet"
      Case 26 : Return "Mostly Sunny"
      Case 27 : Return "Chance of Flurries"
      Case 28 : Return "Chance of Sleet"
      Case 29 : Return "Chance of Snow"
      Case 30 : Return "Chance of Storms"
      Case 31 : Return "Clear"
      Case 32 : Return "Flurries"
      Case 33 : Return "Hazy"
      Case 34 : Return "Mostly Cloudy"
      Case 35 : Return "Mostly Sunny"
      Case 36 : Return "Sleet"
      Case 37 : Return "Unknown"
      Case 38 : Return "Chance of Rain Showers"
      Case 39 : Return "Chance of Snow Showers"
      Case 40 : Return "Snow Showers"
      Case 41 : Return "Rain Showers"
      Case 42 : Return "Chance of Rain Showers"
      Case 43 : Return "Chance of Snow Showers"
      Case 44 : Return "Snow Showers"
      Case 45 : Return "Rain Showers"
      Case 46 : Return "Freezing Rain"
      Case 47 : Return "Freezing Rain"
      Case 48 : Return "Chance of Freezing Rain"
      Case 49 : Return "Chance of Freezing Rain"
      Case 50 : Return "Windy"
      Case 51 : Return "Foggy"
      Case 52 : Return "Scattered Showers"
      Case 53 : Return "Scattered Thunderstorms"
      Case 54 : Return "Light Snow"
      Case 55 : Return "Chance of Light Snow"
      Case 56 : Return "Frozen Mix"
      Case 57 : Return "Chance of Frozen Mix"
      Case 58 : Return "Drizzle"
      Case 59 : Return "Chance of Drizzle"
      Case 60 : Return "Freezing Drizzle"
      Case 61 : Return "Chance of Freezing Drizzle"
      Case 62 : Return "Heavy Snow"
      Case 63 : Return "Heavy Rain"
      Case 64 : Return "Hot and Humid"
      Case 65 : Return "Very Hot"
      Case 66 : Return "Increasing Clouds"
      Case 67 : Return "Clearing"
      Case 68 : Return "Mostly Cloudy"
      Case 69 : Return "Very Cold"
      Case 70 : Return "Mostly Clear"
      Case 71 : Return "Increasing Clouds"
      Case 72 : Return "Clearing"
      Case 73 : Return "Mostly Cloudy"
      Case 74 : Return "Very Cold"
      Case 75 : Return "Warm and Humid"
      Case 76 : Return "Now"
      Case 77 : Return "Exclamation"
      Case 78 : Return "30% Chance of Snow"
      Case 79 : Return "40% Chance of Snow"
      Case 80 : Return "50% Chance of Snow"
      Case 81 : Return "30% Chance of Rain"
      Case 82 : Return "40% Chance of Rain"
      Case 83 : Return "50% Chance of Rain"
      Case 84 : Return "30% Chance of Flurries"
      Case 85 : Return "40% Chance of Flurries"
      Case 86 : Return "50% Chance of Flurries"
      Case 87 : Return "30% Chance of Rain"
      Case 88 : Return "40% Chance of Rain"
      Case 89 : Return "50% Chance of Rain"
      Case 90 : Return "30% Chance of Sleet"
      Case 91 : Return "40% Chance of Sleet"
      Case 92 : Return "50% Chance of Sleet"
      Case 93 : Return "30% Chance of Storms"
      Case 94 : Return "40% Chance of Storms"
      Case 95 : Return "50% Chance of Storms"
      Case 96 : Return "30% Chance of Flurries"
      Case 97 : Return "40% Chance of Flurries"
      Case 98 : Return "50% Chance of Flurries"
      Case 99 : Return "30% Chance of Sleet"
      Case 100 : Return "40% Chance of Sleet"
      Case 101 : Return "50% Chance of Sleet"
      Case 102 : Return "30% Chance of Snow"
      Case 103 : Return "40% Chance of Snow"
      Case 104 : Return "50% Chance of Snow"
      Case 105 : Return "30% Chance of Storms"
      Case 106 : Return "40% Chance of Storms"
      Case 107 : Return "50% Chance of Storms"
      Case 108 : Return "30% Chance Rain Showers"
      Case 109 : Return "40% Chance Rain Showers"
      Case 110 : Return "50% Chance Rain Showers"
      Case 111 : Return "30% Chance Snow Showers"
      Case 112 : Return "40% Chance Snow Showers"
      Case 113 : Return "50% Chance Snow Showers"
      Case 114 : Return "30% Chance Rain Showers"
      Case 115 : Return "40% Chance Rain Showers"
      Case 116 : Return "50% Chance Rain Showers"
      Case 117 : Return "30% Chance Snow Showers"
      Case 118 : Return "40% Chance Snow Showers"
      Case 119 : Return "50% Chance Snow Showers"
      Case 120 : Return "30% Chance Freezing Rain"
      Case 121 : Return "40% Chance Freezing Rain"
      Case 122 : Return "50% Chance Freezing Rain"
      Case 123 : Return "30% Chance Freezing Rain"
      Case 124 : Return "40% Chance Freezing Rain"
      Case 125 : Return "50% Chance Freezing Rain"
      Case 126 : Return "30% Chance Light Snow"
      Case 127 : Return "40% Chance Light Snow"
      Case 128 : Return "50% Chance Light Snow"
      Case 129 : Return "30% Chance Frozen Mix"
      Case 130 : Return "40% Chance Frozen Mix"
      Case 131 : Return "50% Chance Frozen Mix"
      Case 132 : Return "30% Chance of Drizzle"
      Case 133 : Return "40% Chance of Drizzle"
      Case 134 : Return "50% Chance of Drizzle"
      Case 135 : Return "30% Chance Freezing Drizzle"
      Case 136 : Return "40% Chance Freezing Drizzle"
      Case 137 : Return "50% Chance Freezing Drizzle"
      Case 138 : Return "Chance of Snow"
      Case 139 : Return "Chance of Rain"
      Case 140 : Return "Chance of Flurries"
      Case 141 : Return "Chance of Rain"
      Case 142 : Return "Chance of Sleet"
      Case 143 : Return "Chance of Storms"
      Case 144 : Return "Chance of Flurries"
      Case 145 : Return "Chance of Sleet"
      Case 146 : Return "Chance of Snow"
      Case 147 : Return "Chance of Storms"
      Case 148 : Return "Chance of Rain Showers"
      Case 149 : Return "Chance of Snow Showers"
      Case 150 : Return "Chance of Rain Showers"
      Case 151 : Return "Chance of Snow Showers"
      Case 152 : Return "Chance of Freezing Rain"
      Case 153 : Return "Chance of Freezing Rain"
      Case 154 : Return "Chance of Light Snow"
      Case 155 : Return "Chance of Frozen Mix"
      Case 156 : Return "Chance of Drizzle"
      Case 157 : Return "Chance of Freezing Drizzle"
      Case 158 : Return "Windy"
      Case 159 : Return "Foggy"
      Case 160 : Return "Light Snow"
      Case 161 : Return "Frozen Mix"
      Case 162 : Return "Drizzle"
      Case 163 : Return "Heavy Rain"
      Case 164 : Return "Chance of Frozen Mix"
      Case 165 : Return "Chance of Drizzle"
      Case 166 : Return "Chance of Frozen Drizzle"
      Case 167 : Return "30% Chance of Drizzle"
      Case 168 : Return "30% Chance Frozen Drizzle"
      Case 169 : Return "30% Chance Frozen Mix"
      Case 170 : Return "40% Chance of Drizzle"
      Case 171 : Return "40% Chance Frozen Drizzle"
      Case 172 : Return "40% Chance Frozen Mix"
      Case 173 : Return "50% Chance of Drizzle"
      Case 174 : Return "50% Chance Frozen Drizzle"
      Case 175 : Return "50% Chance Frozen Mix"
      Case 176 : Return "Chance of Light Snow"
      Case 177 : Return "30% Chance Light Snow"
      Case 178 : Return "40% Chance Light Snow"
      Case 179 : Return "50% Chance Light Snow"
      Case 180 : Return "Scattered Thunderstorms"
      Case 181 : Return "Freezing Drizzle"
      Case 182 : Return "Scattered Showers"
      Case 183 : Return "Scattered Thunderstorms"
      Case 184 : Return "Warm and Humid"
      Case 185 : Return "60% Chance of Snow"
      Case 186 : Return "70% Chance of Snow"
      Case 187 : Return "80% Chance of Snow"
      Case 188 : Return "60% Chance of Rain"
      Case 189 : Return "70% Chance of Rain"
      Case 190 : Return "80% Chance of Rain"
      Case 191 : Return "60% Chance of Flurries"
      Case 192 : Return "70% Chance of Flurries"
      Case 193 : Return "80% Chance of Flurries"
      Case 194 : Return "60% Chance of Rain"
      Case 195 : Return "70% Chance of Rain"
      Case 196 : Return "80% Chance of Rain"
      Case 197 : Return "60% Chance of Sleet"
      Case 198 : Return "70% Chance of Sleet"
      Case 199 : Return "80% Chance of Sleet"
      Case 200 : Return "60% Chance of Storms"
      Case 201 : Return "70% Chance of Storms"
      Case 202 : Return "80% Chance of Storms"
      Case 203 : Return "60% Chance of Flurries"
      Case 204 : Return "70% Chance of Flurries"
      Case 205 : Return "80% Chance of Flurries"
      Case 206 : Return "60% Chance of Sleet"
      Case 207 : Return "70% Chance of Sleet"
      Case 208 : Return "80% Chance of Sleet"
      Case 209 : Return "60% Chance of Snow"
      Case 210 : Return "70% Chance of Snow"
      Case 211 : Return "80% Chance of Snow"
      Case 212 : Return "60% Chance of Storms"
      Case 213 : Return "70% Chance of Storms"
      Case 214 : Return "80% Chance of Storms"
      Case 215 : Return "60% Chance Rain Showers"
      Case 216 : Return "70% Chance Rain Showers"
      Case 217 : Return "80% Chance Rain Showers"
      Case 218 : Return "60% Chance Snow Showers"
      Case 219 : Return "70% Chance Snow Showers"
      Case 220 : Return "80% Chance Snow Showers"
      Case 221 : Return "60% Chance Rain Showers"
      Case 222 : Return "70% Chance Rain Showers"
      Case 223 : Return "80% Chance Rain Showers"
      Case 224 : Return "60% Chance Snow Showers"
      Case 225 : Return "70% Chance Snow Showers"
      Case 226 : Return "80% Chance Snow Showers"
      Case 227 : Return "60% Chance Freezing Rain"
      Case 228 : Return "70% Chance  Freezing Rain"
      Case 229 : Return "80% Chance Freezing Rain"
      Case 230 : Return "60% Chance Freezing Rain"
      Case 231 : Return "70% Chance Freezing Rain"
      Case 232 : Return "80% Chance Freezing Rain"
      Case 233 : Return "60% Chance Light Snow"
      Case 234 : Return "70% Chance Light Snow"
      Case 235 : Return "80% Chance Light Snow"
      Case 236 : Return "60% Chance Frozen Mix"
      Case 236 : Return "70% Chance Frozen Mix"
      Case 238 : Return "80% Chance Frozen Mix"
      Case 239 : Return "60% Chance of Drizzle"
      Case 240 : Return "70% Chance of Drizzle"
      Case 241 : Return "80% Chance of Drizzle"
      Case 242 : Return "60% Chance Freezing Drizzle"
      Case 243 : Return "70% Chance Freezing Drizzle"
      Case 244 : Return "80% Chance Freezing Drizzle"
      Case 245 : Return "60% Chance Light Snow"
      Case 246 : Return "70% Chance Light Snow"
      Case 247 : Return "80% Chance Light Snow"
      Case 248 : Return "60% Chance Frozen Mix"
      Case 249 : Return "70% Chance Frozen Mix"
      Case 250 : Return "80% Chance Frozen Mix"
      Case 251 : Return "Chance of Light Rain"
      Case 252 : Return "30% Chance of Light Rain"
      Case 253 : Return "40% Chance of Light Rain"
      Case 254 : Return "50% Chance of Light Rain"
      Case 255 : Return "60% Chance of Light Rain"
      Case 256 : Return "70% Chance of Light Rain"
      Case 257 : Return "80% Chance of Light Rain"
      Case 258 : Return "Light Rain"
      Case 259 : Return "Chance of Light Rain"
      Case 260 : Return "30% Chance of Light Rain"
      Case 261 : Return "40% Chance of Light Rain"
      Case 262 : Return "50% Chance of Light Rain"
      Case 263 : Return "60% Chance of Light Rain"
      Case 264 : Return "70% Chance of Light Rain"
      Case 265 : Return "80% Chance of Light Rain"
      Case 266 : Return "Light Rain"
      Case 267 : Return "Heavy Snow"
      Case 268 : Return "Chance of Heavy Snow"
      Case 269 : Return "30 % Chance of Heavy Snow"
      Case 270 : Return "40% Chance of Heavy Snow"
      Case 271 : Return "50% Chance of Heavy Snow"
      Case 272 : Return "60% Chance of Heavy Snow"
      Case 273 : Return "70% Chance of Heavy Snow"
      Case 274 : Return "80% Chance of Heavy Snow"
      Case 275 : Return "Heavy Snow"
      Case 276 : Return "Chance of Heavy Snow"
      Case 277 : Return "30% Chance of Heavy Snow"
      Case 278 : Return "40% Chance of Heavy Snow"
      Case 279 : Return "50% Chance of Heavy Snow"
      Case 280 : Return "60% Chance of Heavy Snow"
      Case 281 : Return "70% Chance of Heavy Snow"
      Case 282 : Return "80% Chance of Heavy Snow"
      Case Else : Return "Unknown"
    End Select

  End Function

  ''' <summary>
  ''' Returns the Condition Category Name
  ''' </summary>
  ''' <param name="iCondition"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWeatherConditionNameByValue(ByVal iCondition As Integer) As String

    Select Case iCondition
      Case 0 : Return "Unknown"
      Case 1 : Return "Sunny/Clear"
      Case 2 : Return "Cloudy/Fair/Hazy"
      Case 3 : Return "Rain/Rain Showers/Drizzle"
      Case 4 : Return "Sleet/Freezing Rain/Freezing Drizzle"
      Case 5 : Return "Snow/Snow Showers/Flurries"
      Case 6 : Return "Storms/Thunderstorms"
      Case 7 : Return "Foggy"
      Case 8 : Return "Windy"
      Case 9 : Return "Warm and Humid"
      Case 10 : Return "Cold and Dry"
      Case Else : Return "Unknown"
    End Select

  End Function

  ''' <summary>
  ''' Returns the Condition Category Id
  ''' </summary>
  ''' <param name="strConditionName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWeatherConditionValueByName(ByVal strConditionName As String) As Integer

    Select Case strConditionName
      Case "Unknown" : Return 0
      Case "Sunny/Clear" : Return 1
      Case "Cloudy/Fair/Hazy" : Return 2
      Case "Rain/Rain Showers/Drizzle" : Return 3
      Case "Sleet/Freezing Rain/Freezing Drizzle" : Return 4
      Case "Snow/Snow Showers/Flurries" : Return 5
      Case "Storms/Thunderstorms" : Return 6
      Case "Foggy" : Return 7
      Case "Windy" : Return 8
      Case "Warm and Humid" : Return 9
      Case "Cold and Dry" : Return 10
      Case Else : Return 0
    End Select

  End Function

  Public Function GetWeatherConditionValueByRegex(ByVal strConditionName As String) As Integer

    If Regex.IsMatch(strConditionName, "Unknown", RegexOptions.IgnoreCase) = True Then Return 0
    If Regex.IsMatch(strConditionName, "Sunny|Clear", RegexOptions.IgnoreCase) = True Then Return 1
    If Regex.IsMatch(strConditionName, "Cloudy|Fair|Hazy", RegexOptions.IgnoreCase) = True Then Return 2
    If Regex.IsMatch(strConditionName, "Rain|Rain Showers|Drizzle", RegexOptions.IgnoreCase) = True Then Return 3
    If Regex.IsMatch(strConditionName, "Sleet|Freezing Rain|Freezing Drizzle|Frozen Mix|Frozen Drizzle", RegexOptions.IgnoreCase) = True Then Return 4
    If Regex.IsMatch(strConditionName, "Snow|Snow Showers|Flurries", RegexOptions.IgnoreCase) = True Then Return 5
    If Regex.IsMatch(strConditionName, "Storms|Thunderstorms", RegexOptions.IgnoreCase) = True Then Return 6
    If Regex.IsMatch(strConditionName, "Foggy", RegexOptions.IgnoreCase) = True Then Return 7
    If Regex.IsMatch(strConditionName, "Windy", RegexOptions.IgnoreCase) = True Then Return 8
    If Regex.IsMatch(strConditionName, "Warm and Humid", RegexOptions.IgnoreCase) = True Then Return 9
    If Regex.IsMatch(strConditionName, "Cold and Dry", RegexOptions.IgnoreCase) = True Then Return 10
    Return 0

  End Function

  ''' <summary>
  ''' Returns the Weather Condition Category by Condition Id
  ''' </summary>
  ''' <param name="strCondition"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWeatherConditionCategory(ByVal strCondition As String) As Integer

    Select Case strCondition
      Case "999" : Return 0 ' Unknown
      Case "000" : Return 1 ' Clear
      Case "001" : Return 2 ' Cloudy
      Case "002" : Return 2 ' Partly Cloudy
      Case "003" : Return 2 ' Partly Cloudy
      Case "004" : Return 1 ' Partly Sunny
      Case "005" : Return 3 ' Rain
      Case "006" : Return 6 ' Thunderstorms
      Case "007" : Return 1 ' Sunny
      Case "008" : Return 5 ' Snow
      Case "009" : Return 5 ' Flurries
      Case "010" : Return 0 ' Unknown
      Case "011" : Return 5 ' Chance of Snow
      Case "012" : Return 5 ' Snow
      Case "013" : Return 2 ' Cloudy
      Case "014" : Return 3 ' Rain
      Case "015" : Return 3 ' Chance of Rain
      Case "016" : Return 2 ' Partly Cloudy
      Case "017" : Return 1 ' Clear
      Case "018" : Return 6 ' Thunderstorms
      Case "019" : Return 5 ' Chance of Flurries
      Case "020" : Return 3 ' Chance of Rain
      Case "021" : Return 4 ' Chance of Sleet
      Case "022" : Return 6 ' Chance of Storms
      Case "023" : Return 2 ' Hazy
      Case "024" : Return 2 ' Mostly Cloudy
      Case "025" : Return 4 ' Sleet
      Case "026" : Return 1 ' Mostly Sunny
      Case "027" : Return 5 ' Chance of Flurries
      Case "028" : Return 4 ' Chance of Sleet
      Case "029" : Return 5 ' Chance of Snow
      Case "030" : Return 6 ' Chance of Storms
      Case "031" : Return 1 ' Clear
      Case "032" : Return 5 ' Flurries
      Case "033" : Return 2 ' Hazy
      Case "034" : Return 2 ' Mostly Cloudy
      Case "035" : Return 1 ' Mostly Sunny
      Case "036" : Return 4 ' Sleet
      Case "037" : Return 0 ' Unknown
      Case "038" : Return 3 ' Chance of Rain Showers
      Case "039" : Return 5 ' Chance of Snow Showers
      Case "040" : Return 5 ' Snow Showers
      Case "041" : Return 3 ' Rain Showers
      Case "042" : Return 3 ' Chance of Rain Showers
      Case "043" : Return 5 ' Chance of Snow Showers
      Case "044" : Return 5 ' Snow Showers
      Case "045" : Return 3 ' Rain Showers
      Case "046" : Return 3 ' Freezing Rain
      Case "047" : Return 3 ' Freezing Rain
      Case "048" : Return 3 ' Chance of Freezing Rain
      Case "049" : Return 3 ' Chance of Freezing Rain
      Case "050" : Return 8 ' Windy
      Case "051" : Return 7 ' Foggy
      Case "052" : Return 0 ' Scattered Showers
      Case "053" : Return 6 ' Scattered Thunderstorms
      Case "054" : Return 5 ' Light Snow
      Case "055" : Return 5 ' Chance of Light Snow
      Case "056" : Return 4 ' Frozen Mix
      Case "057" : Return 4 ' Chance of Frozen Mix
      Case "058" : Return 3 ' Drizzle
      Case "059" : Return 3 ' Chance of Drizzle
      Case "060" : Return 3 ' Freezing Drizzle
      Case "061" : Return 3 ' Chance of Freezing Drizzle
      Case "062" : Return 5 ' Heavy Snow
      Case "063" : Return 3 ' Heavy Rain
      Case "064" : Return 0 ' Hot and Humid
      Case "065" : Return 0 ' Very Hot
      Case "066" : Return 0 ' Increasing Clouds
      Case "067" : Return 1 ' Clearing
      Case "068" : Return 2 ' Mostly Cloudy
      Case "069" : Return 0 ' Very Cold
      Case "070" : Return 1 ' Mostly Clear
      Case "071" : Return 0 ' Increasing Clouds
      Case "072" : Return 1 ' Clearing
      Case "073" : Return 2 ' Mostly Cloudy
      Case "074" : Return 0 ' Very Cold
      Case "075" : Return 9 ' Warm and Humid
      Case "076" : Return 0 ' Now
      Case "077" : Return 0 ' Exclamation
      Case "078" : Return 5 ' 30% Chance of Snow
      Case "079" : Return 5 ' 40% Chance of Snow
      Case "080" : Return 5 ' 50% Chance of Snow
      Case "081" : Return 3 ' 30% Chance of Rain
      Case "082" : Return 3 ' 40% Chance of Rain
      Case "083" : Return 3 ' 50% Chance of Rain
      Case "084" : Return 5 ' 30% Chance of Flurries
      Case "085" : Return 5 ' 40% Chance of Flurries
      Case "086" : Return 5 ' 50% Chance of Flurries
      Case "087" : Return 3 ' 30% Chance of Rain
      Case "088" : Return 3 ' 40% Chance of Rain
      Case "089" : Return 3 ' 50% Chance of Rain
      Case "090" : Return 4 ' 30% Chance of Sleet
      Case "091" : Return 4 ' 40% Chance of Sleet
      Case "092" : Return 4 ' 50% Chance of Sleet
      Case "093" : Return 6 ' 30% Chance of Storms
      Case "094" : Return 6 ' 40% Chance of Storms
      Case "095" : Return 6 ' 50% Chance of Storms
      Case "096" : Return 5 ' 30% Chance of Flurries
      Case "097" : Return 5 ' 40% Chance of Flurries
      Case "098" : Return 5 ' 50% Chance of Flurries
      Case "099" : Return 4 ' 30% Chance of Sleet
      Case "100" : Return 4 ' 40% Chance of Sleet
      Case "101" : Return 4 ' 50% Chance of Sleet
      Case "102" : Return 5 ' 30% Chance of Snow
      Case "103" : Return 5 ' 40% Chance of Snow
      Case "104" : Return 5 ' 50% Chance of Snow
      Case "105" : Return 6 ' 30% Chance of Storms
      Case "106" : Return 6 ' 40% Chance of Storms
      Case "107" : Return 6 ' 50% Chance of Storms
      Case "108" : Return 3 ' 30% Chance Rain Showers
      Case "109" : Return 3 ' 40% Chance Rain Showers
      Case "110" : Return 3 ' 50% Chance Rain Showers
      Case "111" : Return 5 ' 30% Chance Snow Showers
      Case "112" : Return 5 ' 40% Chance Snow Showers
      Case "113" : Return 5 ' 50% Chance Snow Showers
      Case "114" : Return 3 ' 30% Chance Rain Showers
      Case "115" : Return 3 ' 40% Chance Rain Showers
      Case "116" : Return 3 ' 50% Chance Rain Showers
      Case "117" : Return 5 ' 30% Chance Snow Showers
      Case "118" : Return 5 ' 40% Chance Snow Showers
      Case "119" : Return 5 ' 50% Chance Snow Showers
      Case "120" : Return 3 ' 30% Chance Freezing Rain
      Case "121" : Return 3 ' 40% Chance Freezing Rain
      Case "122" : Return 3 ' 50% Chance Freezing Rain
      Case "123" : Return 3 ' 30% Chance Freezing Rain
      Case "124" : Return 3 ' 40% Chance Freezing Rain
      Case "125" : Return 3 ' 50% Chance Freezing Rain
      Case "126" : Return 5 ' 30% Chance Light Snow
      Case "127" : Return 5 ' 40% Chance Light Snow
      Case "128" : Return 5 ' 50% Chance Light Snow
      Case "129" : Return 4 ' 30% Chance Frozen Mix
      Case "130" : Return 4 ' 40% Chance Frozen Mix
      Case "131" : Return 4 ' 50% Chance Frozen Mix
      Case "132" : Return 3 ' 30% Chance of Drizzle
      Case "133" : Return 3 ' 40% Chance of Drizzle
      Case "134" : Return 3 ' 50% Chance of Drizzle
      Case "135" : Return 3 ' 30% Chance Freezing Drizzle
      Case "136" : Return 3 ' 40% Chance Freezing Drizzle
      Case "137" : Return 3 ' 50% Chance Freezing Drizzle
      Case "138" : Return 5 ' Chance of Snow
      Case "139" : Return 3 ' Chance of Rain
      Case "140" : Return 5 ' Chance of Flurries
      Case "141" : Return 3 ' Chance of Rain
      Case "142" : Return 4 ' Chance of Sleet
      Case "143" : Return 6 ' Chance of Storms
      Case "144" : Return 5 ' Chance of Flurries
      Case "145" : Return 4 ' Chance of Sleet
      Case "146" : Return 5 ' Chance of Snow
      Case "147" : Return 6 ' Chance of Storms
      Case "148" : Return 3 ' Chance of Rain Showers
      Case "149" : Return 5 ' Chance of Snow Showers
      Case "150" : Return 3 ' Chance of Rain Showers
      Case "151" : Return 5 ' Chance of Snow Showers
      Case "152" : Return 3 ' Chance of Freezing Rain
      Case "153" : Return 3 ' Chance of Freezing Rain
      Case "154" : Return 5 ' Chance of Light Snow
      Case "155" : Return 4 ' Chance of Frozen Mix
      Case "156" : Return 3 ' Chance of Drizzle
      Case "157" : Return 3 ' Chance of Freezing Drizzle
      Case "158" : Return 8 ' Windy
      Case "159" : Return 7 ' Foggy
      Case "160" : Return 5 ' Light Snow
      Case "161" : Return 4 ' Frozen Mix
      Case "162" : Return 3 ' Drizzle
      Case "163" : Return 3 ' Heavy Rain
      Case "164" : Return 4 ' Chance of Frozen Mix
      Case "165" : Return 3 ' Chance of Drizzle
      Case "166" : Return 4 ' Chance of Frozen Drizzle
      Case "167" : Return 3 ' 30% Chance of Drizzle
      Case "168" : Return 4 ' 30% Chance Frozen Drizzle
      Case "169" : Return 4 ' 30% Chance Frozen Mix
      Case "170" : Return 3 ' 40% Chance of Drizzle
      Case "171" : Return 4 ' 40% Chance Frozen Drizzle
      Case "172" : Return 4 ' 40% Chance Frozen Mix
      Case "173" : Return 3 ' 50% Chance of Drizzle
      Case "174" : Return 4 ' 50% Chance Frozen Drizzle
      Case "175" : Return 4 ' 50% Chance Frozen Mix
      Case "176" : Return 5 ' Chance of Light Snow
      Case "177" : Return 5 ' 30% Chance Light Snow
      Case "178" : Return 5 ' 40% Chance Light Snow
      Case "179" : Return 5 ' 50% Chance Light Snow
      Case "180" : Return 6 ' Scattered Thunderstorms
      Case "181" : Return 3 ' Freezing Drizzle
      Case "182" : Return 0 ' Scattered Showers
      Case "183" : Return 6 ' Scattered Thunderstorms
      Case "184" : Return 9 ' Warm and Humid
      Case "185" : Return 5 ' 60% Chance of Snow
      Case "186" : Return 5 ' 70% Chance of Snow
      Case "187" : Return 5 ' 80% Chance of Snow
      Case "188" : Return 3 ' 60% Chance of Rain
      Case "189" : Return 3 ' 70% Chance of Rain
      Case "190" : Return 3 ' 80% Chance of Rain
      Case "191" : Return 5 ' 60% Chance of Flurries
      Case "192" : Return 5 ' 70% Chance of Flurries
      Case "193" : Return 5 ' 80% Chance of Flurries
      Case "194" : Return 3 ' 60% Chance of Rain
      Case "195" : Return 3 ' 70% Chance of Rain
      Case "196" : Return 3 ' 80% Chance of Rain
      Case "197" : Return 4 ' 60% Chance of Sleet
      Case "198" : Return 4 ' 70% Chance of Sleet
      Case "199" : Return 4 ' 80% Chance of Sleet
      Case "200" : Return 6 ' 60% Chance of Storms
      Case "201" : Return 6 ' 70% Chance of Storms
      Case "202" : Return 6 ' 80% Chance of Storms
      Case "203" : Return 5 ' 60% Chance of Flurries
      Case "204" : Return 5 ' 70% Chance of Flurries
      Case "205" : Return 5 ' 80% Chance of Flurries
      Case "206" : Return 4 ' 60% Chance of Sleet
      Case "207" : Return 4 ' 70% Chance of Sleet
      Case "208" : Return 4 ' 80% Chance of Sleet
      Case "209" : Return 5 ' 60% Chance of Snow
      Case "210" : Return 5 ' 70% Chance of Snow
      Case "211" : Return 5 ' 80% Chance of Snow
      Case "212" : Return 6 ' 60% Chance of Storms
      Case "213" : Return 6 ' 70% Chance of Storms
      Case "214" : Return 6 ' 80% Chance of Storms
      Case "215" : Return 3 ' 60% Chance Rain Showers
      Case "216" : Return 3 ' 70% Chance Rain Showers
      Case "217" : Return 3 ' 80% Chance Rain Showers
      Case "218" : Return 5 ' 60% Chance Snow Showers
      Case "219" : Return 5 ' 70% Chance Snow Showers
      Case "220" : Return 5 ' 80% Chance Snow Showers
      Case "221" : Return 3 ' 60% Chance Rain Showers
      Case "222" : Return 3 ' 70% Chance Rain Showers
      Case "223" : Return 3 ' 80% Chance Rain Showers
      Case "224" : Return 5 ' 60% Chance Snow Showers
      Case "225" : Return 5 ' 70% Chance Snow Showers
      Case "226" : Return 5 ' 80% Chance Snow Showers
      Case "227" : Return 3 ' 60% Chance Freezing Rain
      Case "228" : Return 3 ' 70% Chance  Freezing Rain
      Case "229" : Return 3 ' 80% Chance Freezing Rain
      Case "230" : Return 3 ' 60% Chance Freezing Rain
      Case "231" : Return 3 ' 70% Chance Freezing Rain
      Case "232" : Return 3 ' 80% Chance Freezing Rain
      Case "233" : Return 5 ' 60% Chance Light Snow
      Case "234" : Return 5 ' 70% Chance Light Snow
      Case "235" : Return 5 ' 80% Chance Light Snow
      Case "236" : Return 4 ' 60% Chance Frozen Mix
      Case "237" : Return 0 ' Unknown
      Case "238" : Return 4 ' 80% Chance Frozen Mix
      Case "239" : Return 3 ' 60% Chance of Drizzle
      Case "240" : Return 3 ' 70% Chance of Drizzle
      Case "241" : Return 3 ' 80% Chance of Drizzle
      Case "242" : Return 3 ' 60% Chance Freezing Drizzle
      Case "243" : Return 3 ' 70% Chance Freezing Drizzle
      Case "244" : Return 3 ' 80% Chance Freezing Drizzle
      Case "245" : Return 5 ' 60% Chance Light Snow
      Case "246" : Return 5 ' 70% Chance Light Snow
      Case "247" : Return 5 ' 80% Chance Light Snow
      Case "248" : Return 4 ' 60% Chance Frozen Mix
      Case "249" : Return 4 ' 70% Chance Frozen Mix
      Case "250" : Return 4 ' 80% Chance Frozen Mix
      Case "251" : Return 3 ' Chance of Light Rain
      Case "252" : Return 3 ' 30% Chance of Light Rain
      Case "253" : Return 3 ' 40% Chance of Light Rain
      Case "254" : Return 3 ' 50% Chance of Light Rain
      Case "255" : Return 3 ' 60% Chance of Light Rain
      Case "256" : Return 3 ' 70% Chance of Light Rain
      Case "257" : Return 3 ' 80% Chance of Light Rain
      Case "258" : Return 3 ' Light Rain
      Case "259" : Return 3 ' Chance of Light Rain
      Case "260" : Return 3 ' 30% Chance of Light Rain
      Case "261" : Return 3 ' 40% Chance of Light Rain
      Case "262" : Return 3 ' 50% Chance of Light Rain
      Case "263" : Return 3 ' 60% Chance of Light Rain
      Case "264" : Return 3 ' 70% Chance of Light Rain
      Case "265" : Return 3 ' 80% Chance of Light Rain
      Case "266" : Return 3 ' Light Rain
      Case "267" : Return 5 ' Heavy Snow
      Case "268" : Return 5 ' Chance of Heavy Snow
      Case "269" : Return 5 ' 30 % Chance of Heavy Snow
      Case "270" : Return 5 ' 40% Chance of Heavy Snow
      Case "271" : Return 5 ' 50% Chance of Heavy Snow
      Case "272" : Return 5 ' 60% Chance of Heavy Snow
      Case "273" : Return 5 ' 70% Chance of Heavy Snow
      Case "274" : Return 5 ' 80% Chance of Heavy Snow
      Case "275" : Return 5 ' Heavy Snow
      Case "276" : Return 5 ' Chance of Heavy Snow
      Case "277" : Return 5 ' 30% Chance of Heavy Snow
      Case "278" : Return 5 ' 40% Chance of Heavy Snow
      Case "279" : Return 5 ' 50% Chance of Heavy Snow
      Case "280" : Return 5 ' 60% Chance of Heavy Snow
      Case "281" : Return 5 ' 70% Chance of Heavy Snow
      Case "282" : Return 5 ' 80% Chance of Heavy Snow
      Case Else : Return 0  ' Unknown
    End Select

  End Function

  ''' <summary>
  ''' Refreshes selected Weather Station
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <remarks></remarks>
  Public Sub RefreshStation(ByVal strStationNumber As String)

    Try
      hspi_plugin.GetLiveWeather(strStationNumber)
      Thread.Sleep(1000)
      hspi_plugin.GetForecastWeather(strStationNumber)
    Catch ex As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Search for Station Location
  ''' </summary>
  ''' <param name="searchString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SearchLocations(ByVal searchString As String) As ArrayList

    Try

      Return WeatherBugAPI.GetLocations(searchString)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SearchLocations()")
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Search for Station Ids
  ''' </summary>
  ''' <param name="searchString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SearchStationIds(ByVal searchString As String) As ArrayList

    Try

      Return WeatherBugAPI.GetStationList(searchString)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SearchStationIds()")
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered GetSetting() function.", MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  '''  Saves plug-in settings to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered SaveSetting() subroutine.", MessageType.Debug)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Apply the API Primary Key
      '
      If strSection = "API" And strKey = "KeyPrimary" Then
        gAPIKeyPrimary = strValue
      End If

      '
      ' Apply the API Secret Key
      '
      If strSection = "API" And strKey = "KeySecondary" Then
        gAPIKeySecondary = strValue
      End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "Hashtable Initilization"

  ''' <summary>
  ''' Initialize hash tables used to store weather data
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub InitializeHashTables()

    Try

      Call WriteMessage("Initializing system hash tables ...", MessageType.Debug)

      '
      ' Define Weather Station Devices
      '
      SyncLock Stations.SyncRoot
        Stations.Clear()

        Dim KeyWeatherNames() As String = {"Weather", "Temperature", "Humidity", "Wind", "Rain", "Pressure", "Visibility", "Forecast", "Alerts"}

        For index As Integer = 1 To 5
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)

          For Each strKeyName As String In KeyWeatherNames

            Dim Keys() As String = GetWeatherKeys(strKeyName)

            If Stations.ContainsKey(strStationNumber) = False Then
              Stations.Add(strStationNumber, New Hashtable())
            End If

            For Each strKey As String In Keys
              '
              ' Initialize the hash table
              '
              If Stations(strStationNumber).ContainsKey(strKey) = False Then
                Stations(strStationNumber)(strKey) = New Hashtable()

                Stations(strStationNumber)(strKey)("Name") = GetWeatherName(strKey)
                Stations(strStationNumber)(strKey)("Type") = GetWeatherType(strKey)
                Stations(strStationNumber)(strKey)("Image") = String.Empty
                Stations(strStationNumber)(strKey)("Icon") = String.Empty
                Stations(strStationNumber)(strKey)("Units") = String.Empty
                Stations(strStationNumber)(strKey)("Value") = String.Empty
                Stations(strStationNumber)(strKey)("String") = String.Empty
                Stations(strStationNumber)(strKey)("LastChange") = Now.ToString

                Stations(strStationNumber)(strKey)("DevCode") = GetDeviceAddress(String.Format("{0}-{1}", strStationNumber, strKey))
              End If
            Next
          Next
        Next

      End SyncLock

      '
      ' Define Forecast Devices
      '
      SyncLock Forecasts.SyncRoot
        Forecasts.Clear()

        '
        ' Process each of the 5 supported weather stations
        '
        For index As Integer = 1 To 5
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)

          '
          ' Ensure our hashtable is configured properly
          '
          If Forecasts.ContainsKey(strStationNumber) = False Then
            Forecasts.Add(strStationNumber, New Hashtable())
          End If

          For i As Integer = 1 To 14
            Dim Forecast As New Hashtable

            Forecast.Add("title", String.Empty)
            Forecast.Add("cloudCoverPercent", String.Empty)
            Forecast.Add("dewPoint", String.Empty)
            Forecast.Add("iconCode", 0)
            Forecast.Add("iconImage", String.Empty)
            Forecast.Add("precipCode", String.Empty)
            Forecast.Add("precipProbability", String.Empty)
            Forecast.Add("relativeHumidity", String.Empty)
            Forecast.Add("summaryDescription", String.Empty)
            Forecast.Add("detailedDescription", String.Empty)
            Forecast.Add("temperature", String.Empty)
            Forecast.Add("thunderstormProbability", String.Empty)
            Forecast.Add("windDirectionDegrees", String.Empty)
            Forecast.Add("windSpeed", String.Empty)
            Forecast.Add("forecastDateLocalStr", String.Empty)

            Forecasts(strStationNumber)(i) = Forecast
          Next
        Next

      End SyncLock

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "InitializeHashTables()")
    End Try

    Call WriteMessage("Hash tables intialized.", MessageType.Debug)

  End Sub

  ''' <summary>
  ''' Gets the Weather Friendly Name
  ''' </summary>
  ''' <param name="strKey"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function GetWeatherName(ByVal strKey As String) As String

    Select Case strKey
      Case "ob-date" : Return "Observation Date"
      Case "current-condition" : Return "Current Condition"

      Case "temp" : Return "Temperature"
      Case "temp-rate" : Return "Temperature Rate"
      Case "heat-index" : Return "Heat Index"
      Case "feels-like" : Return "Temperature (Feels Like)"
      Case "dew-point" : Return "Dew Point Temperature"
      Case "dew-point-rate" : Return "Dew Point Temperature Rate"

      Case "humidity" : Return "Humidity"
      Case "humidity-rate" : Return "Humidity Rate"

      Case "gust-direction" : Return "Wind Gust Direction"
      Case "gust-speed" : Return "Gust Speed"
      Case "gust-time" : Return "Wind Gust Time"
      Case "wind-speed" : Return "Wind Speed"
      Case "wind-speed-avg" : Return "Wind Speed Average"
      Case "wind-direction" : Return "Wind Direction"
      Case "wind-direction-avg" : Return "Wind Direction Average"

      Case "rain-month" : Return "Rain Month"
      Case "rain-rate" : Return "Rain Rate"
      Case "rain-today" : Return "Rain Today"
      Case "rain-year" : Return "Rain Year"

      Case "pressure" : Return "Pressure"
      Case "pressure-rate" : Return "Pressure Rate"

      Case "visibility" : Return "Visibility"
      Case "visibility-rate" : Return "Visibility Rate"

      Case "todays-temperature-day" : Return "Today's High"
      Case "todays-temperature-night" : Return "Today's Low"
      Case "todays-short-prediction-day" : Return "Today's Prediction"
      Case "todays-short-prediction-night" : Return "Tonight's Prediction"

      Case "tomorrows-temperature-day" : Return "Tomorrow's High"
      Case "tomorrows-temperature-night" : Return "Tomorrow's Low"
      Case "tomorrows-short-prediction-day" : Return "Tomorrow's Prediction"
      Case "tomorrows-short-prediction-night" : Return "Tomorrow Night's Prediction"

      Case "2-day-temperature-day" : Return "2 Day High"
      Case "2-day-temperature-night" : Return "2 Day Low"
      Case "2-day-short-prediction-day" : Return "2 Day Prediction (Day)"
      Case "2-day-short-prediction-night" : Return "2 Day Prediction (Night)"

      Case "3-day-temperature-day" : Return "3 Day High"
      Case "3-day-temperature-night" : Return "3 Day Low"
      Case "3-day-short-prediction-day" : Return "3 Day Prediction (Day)"
      Case "3-day-short-prediction-night" : Return "3 Day Prediction (Night)"

      Case "4-day-temperature-day" : Return "4 Day High"
      Case "4-day-temperature-night" : Return "4 Day Low"
      Case "4-day-short-prediction-day" : Return "4 Day Prediction (Day)"
      Case "4-day-short-prediction-night" : Return "4 Day Prediction (Night)"

      Case "5-day-temperature-day" : Return "5 Day High"
      Case "5-day-temperature-night" : Return "5 Day Low"
      Case "5-day-short-prediction-day" : Return "5 Day Prediction (Day)"
      Case "5-day-short-prediction-night" : Return "5 Day Prediction (Night)"

      Case "6-day-temperature-day" : Return "6 Day High"
      Case "6-day-temperature-night" : Return "6 Day Low"
      Case "6-day-short-prediction-day" : Return "6 Day Prediction (Day)"
      Case "6-day-short-prediction-night" : Return "6 Day Prediction (Night)"

      Case "7-day-temperature-day" : Return "7 Day High"
      Case "7-day-temperature-night" : Return "7 Day Low"
      Case "7-day-short-prediction-day" : Return "7 Day Prediction (Day)"
      Case "7-day-short-prediction-night" : Return "7 Day Prediction (Night)"

      Case "last-alert-type" : Return "Last Alert Type"
      Case "last-alert-title" : Return "Last Alert Title"

      Case Else : Return strKey
    End Select

  End Function

  ''' <summary>
  ''' Gets the Weather Type
  ''' </summary>
  ''' <param name="strKey"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function GetWeatherType(ByVal strKey As String) As String

    Select Case strKey
      Case "ob-date" : Return "Weather"
      Case "current-condition" : Return "Weather"

      Case "temp" : Return "Temperature"
      Case "temp-rate" : Return "Temperature"
      Case "heat-index" : Return "Temperature"
      Case "feels-like" : Return "Temperature"
      Case "dew-point" : Return "Temperature"
      Case "dew-point-rate" : Return "Temperature"

      Case "humidity" : Return "Humidity"
      Case "humidity-rate" : Return "Humidity"

      Case "gust-direction" : Return "Wind"
      Case "gust-speed" : Return "Wind"
      Case "gust-time" : Return "Wind"
      Case "wind-speed" : Return "Wind"
      Case "wind-speed-avg" : Return "Wind"
      Case "wind-direction" : Return "Wind"
      Case "wind-direction-avg" : Return "Wind"

      Case "rain-month" : Return "Rain"
      Case "rain-rate" : Return "Rain"
      Case "rain-rate-max" : Return "Rain"
      Case "rain-today" : Return "Rain"
      Case "rain-year" : Return "Rain"

      Case "pressure" : Return "Pressure"
      Case "pressure-rate" : Return "Pressure"

      Case "visibility" : Return "Visibility"
      Case "visibility-rate" : Return "Visibility"

      Case "todays-temperature-day" : Return "Forecast"
      Case "todays-temperature-night" : Return "Forecast"
      Case "todays-short-prediction-day" : Return "Forecast"
      Case "todays-short-prediction-night" : Return "Forecast"

      Case "tomorrows-temperature-day" : Return "Forecast"
      Case "tomorrows-temperature-night" : Return "Forecast"
      Case "tomorrows-short-prediction-day" : Return "Forecast"
      Case "tomorrows-short-prediction-night" : Return "Forecast"

      Case "2-day-temperature-day" : Return "Forecast"
      Case "2-day-temperature-night" : Return "Forecast"
      Case "2-day-short-prediction-day" : Return "Forecast"
      Case "2-day-short-prediction-night" : Return "Forecast"

      Case "3-day-temperature-day" : Return "Forecast"
      Case "3-day-temperature-night" : Return "Forecast"
      Case "3-day-short-prediction-day" : Return "Forecast"
      Case "3-day-short-prediction-night" : Return "Forecast"

      Case "4-day-temperature-day" : Return "Forecast"
      Case "4-day-temperature-night" : Return "Forecast"
      Case "4-day-short-prediction-day" : Return "Forecast"
      Case "4-day-short-prediction-night" : Return "Forecast"

      Case "5-day-temperature-day" : Return "Forecast"
      Case "5-day-temperature-night" : Return "Forecast"
      Case "5-day-short-prediction-day" : Return "Forecast"
      Case "5-day-short-prediction-night" : Return "Forecast"

      Case "6-day-temperature-day" : Return "Forecast"
      Case "6-day-temperature-night" : Return "Forecast"
      Case "6-day-short-prediction-day" : Return "Forecast"
      Case "6-day-short-prediction-night" : Return "Forecast"

      Case "7-day-temperature-day" : Return "Forecast"
      Case "7-day-temperature-night" : Return "Forecast"
      Case "7-day-short-prediction-day" : Return "Forecast"
      Case "7-day-short-prediction-night" : Return "Forecast"

      Case "last-alert-type" : Return "Alerts"
      Case "last-alert-title" : Return "Alerts"

      Case Else : Return "Unknown Device"
    End Select

  End Function

  ''' <summary>
  ''' Gets the Key Names
  ''' </summary>
  ''' <param name="strKeyType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetWeatherKeys(ByVal strKeyType As String) As String()

    Try

      ' "Weather", "Temperature", "Humidity", "Wind", "Rain", "Pressure", "Visibility", "Forecast", "Alerts"

      Select Case strKeyType
        Case "Temperature"
          Dim Keys() As String = {"temp", _
                                  "temp-rate", _
                                  "heat-index", _
                                  "feels-like", _
                                  "dew-point", _
                                  "dew-point-rate"}
          Return Keys

        Case "Humidity"
          Dim Keys() As String = {"humidity", _
                                  "humidity-rate"}
          Return Keys

        Case "Wind"
          Dim Keys() As String = {"gust-direction", _
                                  "gust-speed", _
                                  "gust-time", _
                                  "wind-speed", _
                                  "wind-speed-avg", _
                                  "wind-direction", _
                                  "wind-direction-avg"}
          Return Keys

        Case "Rain"
          Dim Keys() As String = {"rain-month", _
                                  "rain-rate", _
                                  "rain-today", _
                                  "rain-year"}
          Return Keys

        Case "Pressure"
          Dim Keys() As String = {"pressure", _
                                  "pressure-rate"}
          Return Keys

        Case "Visibility"
          Dim Keys() As String = {"visibility", _
                                  "visibility-rate"}
          Return Keys

        Case "Weather"
          Dim Keys() As String = {"ob-date", _
                                  "current-condition"}

          Return Keys

        Case "Forecast"

          Dim Keys() As String = {"todays-temperature-day", _
                                  "todays-temperature-night", _
                                  "todays-short-prediction-day", _
                                  "todays-short-prediction-night", _
                                  "tomorrows-temperature-day", _
                                  "tomorrows-temperature-night", _
                                  "tomorrows-short-prediction-day", _
                                  "tomorrows-short-prediction-night", _
                                  "2-day-temperature-day", _
                                  "2-day-temperature-night", _
                                  "2-day-short-prediction-day", _
                                  "2-day-short-prediction-night", _
                                  "3-day-temperature-day", _
                                  "3-day-temperature-night", _
                                  "3-day-short-prediction-day", _
                                  "3-day-short-prediction-night", _
                                  "4-day-temperature-day", _
                                  "4-day-temperature-night", _
                                  "4-day-short-prediction-day", _
                                  "4-day-short-prediction-night", _
                                  "5-day-temperature-day", _
                                  "5-day-temperature-night", _
                                  "5-day-short-prediction-day", _
                                  "5-day-short-prediction-night", _
                                  "6-day-temperature-day", _
                                  "6-day-temperature-night", _
                                  "6-day-short-prediction-day", _
                                  "6-day-short-prediction-night", _
                                  "7-day-temperature-day", _
                                  "7-day-temperature-night", _
                                  "7-day-short-prediction-day", _
                                  "7-day-short-prediction-night"}

          Return Keys

        Case "Alerts"
          Dim Keys() As String = {"last-alert-type", _
                                  "last-alert-title"}
          Return Keys

        Case Else
          Return Nothing
      End Select

    Catch ex As Exception
      Return Nothing
    End Try

  End Function

#End Region

#Region "UltraNetCam3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      triggers.Add(o, "Weather Alert")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      'Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber
      Case WeatherTriggers.WeatherAlert
        Dim triggerName As String = GetEnumName(WeatherTriggers.WeatherAlert)

        '
        ' Start Alert Type
        '
        Dim ActionSelected As String = trigger.Item("AlertType")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "AlertType", UID, sUnique)

        Dim jqAlertType As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqAlertType.autoPostBack = True

        jqAlertType.AddItem("(Select Alert Type)", "", (ActionSelected = ""))
        Dim names As String() = System.Enum.GetNames(GetType(WeatherAlertTypes))
        For i As Integer = 0 To names.Length - 1
          Dim strOptionName = names(i)
          Dim strOptionValue = names(i)
          jqAlertType.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqAlertType.Build)

        '
        ' Start Station Name
        '
        ActionSelected = trigger.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", triggerName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("reported by")
        stb.Append(jqStation.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case WeatherTriggers.WeatherAlert
          Dim triggerName As String = GetEnumName(WeatherTriggers.WeatherAlert)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "AlertType_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("AlertType") = ActionValue

              Case InStr(sKey, triggerName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("Station") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case WeatherTriggers.WeatherAlert
          If trigger.Item("AlertType") = "" Then Configured = False
          If trigger.Item("Station") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case WeatherTriggers.WeatherAlert
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = "Weather Alert"
          Dim strAlertType As String = trigger.Item("AlertType")

          Dim strStationNumber As String = trigger.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} on {2}", strAlertType, strTriggerName, strStationNumber)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Private Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case WeatherTriggers.WeatherAlert
                Dim strTriggerName As String = "Weather Alert Trigger"
                Dim strAlertType As String = trigger.Item("AlertType")
                Dim strStationNumber As String = trigger.Item("Station")

                Dim strTriggerCheck As String = String.Format("{0},{1},{2}", strTriggerName, strStationNumber, strAlertType)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      actions.Add(o, "Email Notification")          ' 1
      actions.Add(o, "Speak Weather")               ' 2
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = ActInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case WeatherActions.EmailNotification
        Dim actionName As String = GetEnumName(WeatherActions.EmailNotification)

        '
        ' Start EmailNotification
        '
        Dim ActionSelected As String = action.Item("Notification")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "Notification", UID, sUnique)

        Dim jqNotifications As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNotifications.autoPostBack = True

        jqNotifications.AddItem("(Select E-mail Notification)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Weather Conditions", "Weather Forecast", "Weather Alerts"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqNotifications.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqNotifications.Build)

        '
        ' Start Station Name
        '
        ActionSelected = action.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("from")
        stb.Append(jqStation.Build)

      Case WeatherActions.SpeakWeather
        Dim actionName As String = GetEnumName(WeatherActions.SpeakWeather)

        '
        ' Start Speak Weather
        '
        Dim ActionSelected As String = action.Item("Notification")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "Notification", UID, sUnique)

        Dim jqNotifications As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqNotifications.autoPostBack = True

        jqNotifications.AddItem("(Select Speak Action)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Weather Conditions", "Weather Forecast", "Weather Alerts"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqNotifications.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqNotifications.Build)

        '
        ' Start Station Name
        '
        ActionSelected = action.Item("Station")

        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "Station", UID, sUnique)

        Dim jqStation As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqStation.autoPostBack = True

        jqStation.AddItem("(Select Weather Station)", "", (ActionSelected = ""))

        For index As Integer = 1 To 6
          Dim strStationNumber As String = String.Format("Station{0}", index.ToString)
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", "Not defined", gINIFile))

          Dim strOptionValue As String = strStationNumber
          Dim strOptionName As String = String.Format("{0} [{1}]", strStationNumber, strStationName)
          jqStation.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append("from")
        stb.Append(jqStation.Build)

        '
        ' Start Speaker Host
        '
        ActionSelected = IIf(action.Item("SpeakerHost").Length = 0, "*:*", action.Item("SpeakerHost"))
        actionId = String.Format("{0}{1}_{2}_{3}", actionName, "SpeakerHost", UID, sUnique)

        Dim jqSpeakerHost As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 45, True)
        stb.Append("Host(host:instance)")
        stb.Append(jqSpeakerHost.Build)

    End Select

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case WeatherActions.EmailNotification
          Dim actionName As String = GetEnumName(WeatherActions.EmailNotification)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "Notification_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Notification") = ActionValue

              Case InStr(sKey, actionName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Station") = ActionValue

            End Select
          Next

        Case WeatherActions.SpeakWeather
          Dim actionName As String = GetEnumName(WeatherActions.SpeakWeather)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "Notification_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Notification") = ActionValue

              Case InStr(sKey, actionName & "Station_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("Station") = ActionValue

              Case InStr(sKey, actionName & "SpeakerHost_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("SpeakerHost") = ActionValue
            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case WeatherActions.EmailNotification
        If action.Item("Notification") = "" Then Configured = False
        If action.Item("Station") = "" Then Configured = False

      Case WeatherActions.SpeakWeather
        If action.Item("Notification") = "" Then Configured = False
        If action.Item("Station") = "" Then Configured = False
        If action.Item("SpeakerHost") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber

      Case WeatherActions.EmailNotification
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(WeatherActions.EmailNotification)

          Dim strNotificationType As String = action.Item("Notification")

          Dim strStationNumber As String = action.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} {2}", strActionName, strStationNumber, strNotificationType)
        End If

      Case WeatherActions.SpeakWeather
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(WeatherActions.SpeakWeather)

          Dim strNotificationType As String = action.Item("Notification")

          Dim strStationNumber As String = action.Item("Station")
          Dim strStationName As String = Trim(hs.GetINISetting(strStationNumber, "StationName", strStationNumber, gINIFile))

          Dim strSpeakerHost As String = action.Item("SpeakerHost")

          If strStationName.Length > 0 Then
            strStationNumber = String.Format("{0} [{1}]", strStationNumber, strStationName)
          End If

          stb.AppendFormat("{0} {1} from {2} on {3}", strActionName, strNotificationType, strStationNumber, strSpeakerHost)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber

        Case WeatherActions.EmailNotification
          Dim strNotificationType As String = action.Item("Notification")
          Dim strStationNumber As String = action.Item("Station")

          Select Case strNotificationType
            Case "Weather Conditions"
              EmailWeatherConditions(strStationNumber)

            Case "Weather Forecast"
              EmailWeatherForecast(strStationNumber)

            Case "Weather Alerts"
              EmailWeatherAlerts(strStationNumber)

          End Select

        Case WeatherActions.SpeakWeather
          Dim strNotificationType As String = action.Item("Notification")
          Dim strStationNumber As String = action.Item("Station")
          Dim strSpeakerHost As String = action.Item("SpeakerHost")

          Select Case strNotificationType
            Case "Weather Conditions"
              SpeakWeatherConditions(strStationNumber, False, strSpeakerHost)

            Case "Weather Forecast"
              SpeakWeatherForecast(strStationNumber, False, strSpeakerHost)

            Case "Weather Alerts"
              SpeakWeatherAlerts(strStationNumber, False, strSpeakerHost)
          End Select

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum WeatherTriggers
  <Description("Weather Alert")> _
  WeatherAlert = 1
End Enum

Public Enum WeatherActions
  <Description("Email Notification")> _
  EmailNotification = 1
  <Description("Speak Weather")> _
  SpeakWeather = 2
End Enum

<Flags()> Public Enum WeatherAlertTypes
  Any = 0
  Forecast = 2
  Statement = 4
  Synopsis = 8
  Outlook = 16
  Watch = 32
  Advisory = 64
  Warning = 128
End Enum

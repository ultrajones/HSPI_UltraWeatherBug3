Imports System.Text
Imports System.Web
Imports Scheduler
Imports HomeSeerAPI
Imports System.Collections.Specialized
Imports System.Web.UI.WebControls

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder

      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)
      End If

      Dim Header As New StringBuilder
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultraweatherbug3/js/hspi_ultraweatherbug3_utility.js""></script>")
      Header.AppendLine("<link type=""text/css"" rel=""stylesheet"" href=""/hspi_ultraweatherbug3/css/hspi_ultraweatherbug3.css"" />")
      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      Me.RefreshIntervalMilliSeconds = 3000
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Templates"
      tab.tabDIVID = "tabTemplates"
      tab.tabContent = "<div id='divTemplates'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Stations"
      tab.tabDIVID = "tabStations"
      tab.tabContent = "<div id='divStations'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Devices"
      tab.tabDIVID = "tabDevices"
      tab.tabContent = "<div id='divDevices'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Live Weather"
      tab.tabDIVID = "tabLiveWeather"
      tab.tabContent = "<div id='divLiveWeather'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Forecast"
      tab.tabDIVID = "tabForecast"
      tab.tabContent = "<div id='divForecast'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Alerts"
      tab.tabDIVID = "tabAlerts"
      tab.tabContent = "<div id='divAlerts'></div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("     <table style=""width: 100%"">")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("      </tr>")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("      </tr>")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("      </tr>")
      stb.AppendLine("     </table>")
      stb.AppendLine("    </legend>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> WeatherBug Status </legend>")
      stb.AppendLine("     <table style=""width: 100%"">")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Stations:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", hspi_plugin.GetStatistics("StationCount"))
      stb.AppendLine("      </tr>")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Alerts:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", hspi_plugin.GetStatistics("AlertsCount"))
      stb.AppendLine("      </tr>")
      stb.AppendLine("     </table>")
      stb.AppendLine("    </legend>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> WeatherBug Pulse API </legend>")
      stb.AppendLine("     <table style=""width: 100%"">")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Valid&nbsp;Key:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", WeatherBugAPI.ValidAPIKey())
      stb.AppendLine("      </tr>")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Success:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", hspi_plugin.GetStatistics("APISuccess"))
      stb.AppendLine("      </tr>")
      stb.AppendLine("      <tr>")
      stb.AppendLine("       <td style=""width: 20%""><strong>Failure:</strong></td>")
      stb.AppendFormat("     <td style=""text-align: right"">{0}</td>", hspi_plugin.GetStatistics("APIFailure"))
      stb.AppendLine("      </tr>")
      stb.AppendLine("     </table>")
      stb.AppendLine("    </legend>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      '
      ' WeatherBug API Key
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>WeatherBug API Key</td>")
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug API Key (Primary Key)
      '
      Dim txtAPIKeyPrimary As String = GetSetting("API", "KeyPrimary", "")
      Dim tbAPIKeyPrimary As New clsJQuery.jqTextBox("txtAPIKeyPrimary", "text", txtAPIKeyPrimary, PageName, 40, False)
      tbAPIKeyPrimary.id = "txtAPIKeyPrimary"
      tbAPIKeyPrimary.promptText = "Enter your Pulse API Primary Key issued by WeatherBug.  See the UltraWeatherBug3 HSPI User's Guide for more information."
      tbAPIKeyPrimary.toolTip = tbAPIKeyPrimary.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Primary&nbsp;Key</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbAPIKeyPrimary.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug API Key
      '
      Dim txtAPIKeySecondary As String = GetSetting("API", "KeySecondary", "")
      Dim tbAPIKeySecondary As New clsJQuery.jqTextBox("txtAPIKeySecondary", "text", txtAPIKeySecondary, PageName, 40, False)
      tbAPIKeySecondary.id = "txtAPIKeySecondary"
      tbAPIKeySecondary.promptText = "Enter your Pulse API Secondary Key issued by WeatherBug."
      tbAPIKeySecondary.toolTip = tbAPIKeySecondary.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Secondary&nbsp;Key</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbAPIKeySecondary.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>WeatherBug Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug Options (Unit Type)
      '
      Dim selUnitType As New clsJQuery.jqDropList("selUnitType", Me.PageName, False)
      selUnitType.id = "selUnitType"
      selUnitType.toolTip = "The format used to display temperatures, rainfall and barometric pressure.  The default format is U.S customary units."

      Dim strUnitType As String = GetSetting("Options", "UnitType", "0")
      selUnitType.AddItem("U.S. customary units (miles, °F, etc...)", "0", strUnitType = "0")
      selUnitType.AddItem("Metric system units (kms, °C, etc...)", "1", strUnitType = "1")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Unit&nbsp;Type</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selUnitType.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug Update Frequency
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>WeatherBug Update Frequency</td>")
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug Update Frequency (Station Update)
      '
      Dim selLiveWeatherUpdate As New clsJQuery.jqDropList("selLiveWeatherUpdate", Me.PageName, False)
      selLiveWeatherUpdate.id = "selLiveWeatherUpdate"
      selLiveWeatherUpdate.toolTip = "Specify how often to check the WeatherBug tracking station for updated live weather data."

      Dim txtUpdateFrequency As String = GetSetting("Options", "LiveWeatherUpdate", "5")
      For index As Integer = 5 To 60 Step 5
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Minutes", index.ToString)
        selLiveWeatherUpdate.AddItem(desc, value, index.ToString = txtUpdateFrequency)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Station&nbsp;Update</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLiveWeatherUpdate.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug Update Frequency (Forecast Update)
      '
      Dim selForecastUpdate As New clsJQuery.jqDropList("selForecastUpdate", Me.PageName, False)
      selForecastUpdate.id = "selForecastUpdate"
      selForecastUpdate.toolTip = "Specify how often to check for the weather forecast for the defined tracking stations."

      Dim txtForecastUpdate As String = GetSetting("Options", "ForecastUpdate", "5")
      For index As Integer = 5 To 60 Step 5
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Minutes", index.ToString)
        selForecastUpdate.AddItem(desc, value, index.ToString = txtForecastUpdate)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Forecast&nbsp;Update</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selForecastUpdate.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' WeatherBug Update Frequency (Alerts Update)
      '
      Dim selAlertsUpdate As New clsJQuery.jqDropList("selAlertsUpdate", Me.PageName, False)
      selAlertsUpdate.id = "selAlertsUpdate"
      selAlertsUpdate.toolTip = "Specify how often the plug-in should check for Weather Alerts."

      Dim txtAlertsUpdate As String = GetSetting("Options", "AlertsUpdate", "5")
      For index As Integer = 5 To 60 Step 5
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0} Minutes", index.ToString)
        selAlertsUpdate.AddItem(desc, value, index.ToString = txtAlertsUpdate)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Alerts&nbsp;Update</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selAlertsUpdate.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' HomeSeer Device Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>HomeSeer Device Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' HomeSeer Device Options (Device Image)
      '
      Dim selDeviceImage As New clsJQuery.jqDropList("selDeviceImage", Me.PageName, False)
      selDeviceImage.id = "selDeviceImage"
      selDeviceImage.toolTip = "Display HomeSeer device images."

      Dim txtDeviceImage As String = GetSetting("Options", "DeviceImage", "True")
      selDeviceImage.AddItem("Yes", "True", txtDeviceImage = "True")
      selDeviceImage.AddItem("No", "False", txtDeviceImage = "False")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Device&nbsp;Image</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selDeviceImage.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>E-Mail Notification Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email To)
      '
      Dim txtEmailRcptTo As String = GetSetting("EmailNotification", "EmailRcptTo", "")
      Dim tbEmailRcptTo As New clsJQuery.jqTextBox("txtEmailRcptTo", "text", txtEmailRcptTo, PageName, 60, False)
      tbEmailRcptTo.id = "txtEmailRcptTo"
      tbEmailRcptTo.promptText = "Enter the recipient e-mail address."
      tbEmailRcptTo.toolTip = tbEmailRcptTo.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email To</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailRcptTo.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email From)
      '
      Dim txtEmailFrom As String = GetSetting("EmailNotification", "EmailFrom", "")
      Dim tbEmailFrom As New clsJQuery.jqTextBox("txtEmailFrom", "text", txtEmailFrom, PageName, 60, False)
      tbEmailFrom.id = "txtEmailFrom"
      tbEmailFrom.promptText = "Enter the sender e-mail address."
      tbEmailFrom.toolTip = tbEmailFrom.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email From</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailFrom.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Email Subject)
      '
      Dim txtEmailSubject As String = GetSetting("EmailNotification", "EmailSubject", EMAIL_SUBJECT)
      Dim tbEmailSubject As New clsJQuery.jqTextBox("txtEmailSubject", "text", txtEmailSubject, PageName, 60, False)
      tbEmailSubject.id = "txtEmailSubject"
      tbEmailSubject.promptText = "Enter the subject of the e-mail notification."
      tbEmailSubject.toolTip = tbEmailSubject.promptText

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Email Subject</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", tbEmailSubject.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Web Page Access (Authorized User Roles)
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Options (Logging Level)
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Templates Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabTemplates(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmTemplates", "frmTemplates", "Post"))

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      '
      ' E-Mail Templates
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Email Notification Templates</td>")
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Templates (Live Weather Template)
      '
      Dim chkWeatherCurrent As New clsJQuery.jqCheckBox("chkWeatherCurrent", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkWeatherCurrent.checked = False

      Dim strWeatherTemplate As String = GetSetting("EmailNotification", "WeatherCurrent", WEATHER_CURRENT_TEMPLATE)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Current Weather</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea rows='5' cols='50' name='txtWeatherCurrent'>{0}</textarea>{1}</td>{2}", strWeatherTemplate.Trim.Replace("~", vbCrLf), _
                                                                                                                                  chkWeatherCurrent.Build, _
                                                                                                                                  vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Weather Forecast Template)
      '
      Dim chkWeatherForecast As New clsJQuery.jqCheckBox("chkWeatherForecast", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkWeatherForecast.checked = False

      Dim txtWeatherForecast As String = GetSetting("EmailNotification", "WeatherForecast", WEATHER_FORECAST_TEMPLATE)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Weather Forecasst</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea rows='5' cols='50' name='txtWeatherForecast'>{0}</textarea>{1}</td>{2}", txtWeatherForecast.Trim.Replace("~", vbCrLf), _
                                                                                                                                   chkWeatherForecast.Build, _
                                                                                                                                   vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Weather Alert Template)
      '
      Dim chkWeatherAlert As New clsJQuery.jqCheckBox("chkWeatherAlert", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkWeatherAlert.checked = False

      Dim txtWeatherAlert As String = GetSetting("EmailNotification", "WeatherAlert", WEATHER_ALERT_TEMPLATE)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Weather Alerts</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea rows='5' cols='50' name='txtWeatherAlert'>{0}</textarea>{1}</td>{2}", txtWeatherAlert.Trim.Replace("~", vbCrLf), _
                                                                                                                                chkWeatherAlert.Build, _
                                                                                                                                vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Templates
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Weather To Speech Templates</td>")
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Templates (Live Weather Template)
      '
      Dim chkSpeakWeatherCurrent As New clsJQuery.jqCheckBox("chkSpeakWeatherCurrent", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkSpeakWeatherCurrent.checked = False

      Dim txtSpeakWeatherTemplate As String = GetSetting("Speak", "WeatherCurrent", SPEAK_WEATHER_CURRENT)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Current Weather</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea rows='5' cols='50' name='txtSpeakWeatherTemplate'>{0}</textarea>{1}</td>{2}", txtSpeakWeatherTemplate.Trim.Replace("~", vbCrLf), _
                                                                                                                                        chkSpeakWeatherCurrent.Build, _
                                                                                                                                        vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Weather Forecast Template)
      '
      Dim chkSpeakWeatherForecast As New clsJQuery.jqCheckBox("chkSpeakWeatherForecast", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkSpeakWeatherForecast.checked = False

      Dim txtSpeakWeatherForecast As String = GetSetting("Speak", "WeatherForecast", SPEAK_WEATHER_FORECAST)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Weather Forecast</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea rows='5' cols='50' name='txtSpeakWeatherForecast'>{0}</textarea>{1}</td>{2}", txtSpeakWeatherForecast.Trim.Replace("~", vbCrLf), _
                                                                                                                                        chkSpeakWeatherForecast.Build, _
                                                                                                                                        vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' E-Mail Notification Options (Weather Alert Template)
      '
      Dim chkSpeakWeatherAlert As New clsJQuery.jqCheckBox("chkSpeakWeatherAlert", "&nbsp;Reset To Default", Me.PageName, True, False)
      chkSpeakWeatherAlert.checked = False

      Dim txtSpeakWeatherAlert As String = GetSetting("Speak", "WeatherAlert", SPEAK_WEATHER_ALERT)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Weather Alerts</td>")
      stb.AppendFormat("  <td class='tablecell'><textarea rows='5' cols='50' name='txtSpeakWeatherAlert'>{0}</textarea>{1}</td>{2}", txtSpeakWeatherAlert.Trim.Replace("~", vbCrLf), _
                                                                                                                                     chkSpeakWeatherAlert.Build, _
                                                                                                                                     vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      Dim jqButton1 As New clsJQuery.jqButton("btnSaveTemplates", "Save", Me.PageName, True)
      stb.AppendLine("<div>")
      stb.AppendLine(jqButton1.Build())
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divTemplates", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabTemplates")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Stations Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStations(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmStations", "frmStations", "Post"))

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      '
      ' WeatherBug Stations
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>WeatherBug Stations</td>")
      stb.AppendLine(" </tr>")

      '
      ' Station 1
      '
      For i As Integer = 1 To 5
        Dim strStation As String = String.Format("Station{0}", i.ToString)
        Dim strStationId As String = String.Format("txtStation{0}", i.ToString)
        Dim strButtonId As String = String.Format("btnUpdateStation{0}", i.ToString)
        Dim strButtonClearId As String = String.Format("btnClearStation{0}", i.ToString)

        Dim txtStation As String = String.Format("{0} {1}", GetSetting(strStation, "StationName", ""), GetSetting(strStation, "CityName", ""))
        Dim tbStation As New clsJQuery.jqTextBox(strStationId, "text", txtStation, PageName, 80, False)
        tbStation.id = strStationId
        tbStation.enabled = False

        Dim jqButton1 As New clsJQuery.jqButton(strButtonId, "Select ...", Me.PageName, True)
        Dim jqButton2 As New clsJQuery.jqButton(strButtonClearId, "Clear", Me.PageName, True)
        If i = 1 Then jqButton2.visible = False

        stb.AppendLine(" <tr>")
        stb.AppendFormat("  <td class='tablecell' style=""width: 10%"">Station #{0}</td>{1}", i.ToString, vbCrLf)
        stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;{1}&nbsp;{2}</td>{3}", tbStation.Build, jqButton1.Build(), jqButton2.Build(), vbCrLf)
        stb.AppendLine(" </tr>")
      Next

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStations", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStations")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Station Update Tab
  ''' </summary>
  ''' <param name="StationNumber"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStationUpdate(ByVal StationNumber As String, _
                                 ByVal SearchStage As String, _
                                 ByVal SearchLocation As String, _
                                 ByVal SearchCity As String, _
                                 ByVal SearchStationId As String) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmStationUpdate", "frmStationUpdate", "Post"))

      stb.AppendFormat("<input type='hidden' name='txtStationNumber' value='{0}'>{1}", StationNumber, vbCrLf)

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      '
      ' WeatherBug Stations
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>WeatherBug Station Selection</td>")
      stb.AppendLine(" </tr>")

      Dim tbSearchLocation As New clsJQuery.jqTextBox("txtSearchLocation", "text", SearchLocation, PageName, 80, False)
      tbSearchLocation.id = "txtSearchLocation"
      tbSearchLocation.enabled = True
      tbSearchLocation.editable = True

      Dim jqButton1 As New clsJQuery.jqButton("btnSearchLocation", "Search", Me.PageName, True)
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Search Location</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;{1}</td>{2}", tbSearchLocation.Build(), jqButton1.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Select City
      '
      If SearchLocation.Length > 0 Then

        Dim jqSearchCity As New clsJQuery.jqButton("btnSearchCity", "Select", Me.PageName, True)
        jqSearchCity.enabled = False
        jqSearchCity.visible = False

        Dim selSearchCity As New clsJQuery.jqDropList("selSearchCity", Me.PageName, False)
        selSearchCity.autoPostBack = False
        selSearchCity.id = "selSearchCity"
        selSearchCity.enabled = False
        selSearchCity.visible = False

        Dim locations As ArrayList = hspi_plugin.SearchLocations(SearchLocation)
        If locations.Count = 0 Then
          jqSearchCity.enabled = False
          jqSearchCity.visible = False

          PostMessage(String.Format("No WeatherBug locations were found for {0}.", SearchLocation))
        Else
          jqSearchCity.enabled = True
          jqSearchCity.visible = True

          selSearchCity.toolTip = "Select the WeatherBug City."
          selSearchCity.enabled = True
          selSearchCity.visible = True

          For Each location As Specialized.StringDictionary In locations

            Dim strCityId As String = location("CityId")
            Dim strCityName As String = location("CityName")
            Dim strTerritory As String = location("Territory")  ' State Full Name
            Dim strStateCode As String = location("StateCode")  ' State Code
            Dim strCountry As String = location("Country")

            Dim strLattitude As String = location("Lattitude")
            Dim strLongitude As String = location("Longitude")
            Dim strDma As String = location("Dma")
            Dim strZip As String = location("Zip")

            Dim strLocationValue As String = String.Format("{0},{1}", strLattitude, strLongitude)

            Dim strLocationName As String = ""
            If String.IsNullOrEmpty(strTerritory) Then
              strLocationName = String.Format("{0}, {1}", strCityName, strCountry)
            Else
              strLocationName = String.Format("{0}, {1} {2}", strCityName, strTerritory, strCountry)
            End If

            selSearchCity.AddItem(strLocationName, strLocationValue, strLocationValue = SearchCity)
            If strLocationValue = SearchCity Then
              stb.AppendFormat("<input type='hidden' name='txtCityName' value='{0}'>{1}", strLocationName, vbCrLf)
            End If

          Next

          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Select City</td>")
          stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;{1}</td>{2}", selSearchCity.Build(), jqSearchCity.Build(), vbCrLf)
          stb.AppendLine(" </tr>")

        End If

      End If

      '
      ' Search Station
      '
      If SearchCity.Length > 0 Then

        Dim strButtonCaption As String = IIf(SearchStationId.Length = 0, "Select", "Set Station")
        Dim strButtonId As String = IIf(SearchStationId.Length = 0, "btnSearchStationId", "btnSetStationId")

        Dim jbSearchStationId As New clsJQuery.jqButton(strButtonId, strButtonCaption, Me.PageName, True)
        jbSearchStationId.enabled = False
        jbSearchStationId.visible = False

        Dim selSearchStationId As New clsJQuery.jqDropList("selSearchStationId", Me.PageName, False)
        selSearchStationId.autoPostBack = False
        selSearchStationId.id = "selSearchStationId"
        selSearchStationId.enabled = False
        selSearchStationId.visible = False

        Dim stations As ArrayList = hspi_plugin.SearchStationIds(SearchCity)
        If stations.Count = 0 Then
          jbSearchStationId.enabled = False
          jbSearchStationId.visible = False

          PostMessage(String.Format("No WeatherBug stations were found for {0}.", SearchCity))
        Else

          jbSearchStationId.enabled = True
          jbSearchStationId.visible = True

          selSearchStationId.toolTip = "Select the WeatherBug Station."
          selSearchStationId.enabled = True
          selSearchStationId.visible = True

          For Each station As Specialized.StringDictionary In stations

            Dim strStationId As String = station("StationId")
            Dim strStationName As String = station("StationName")
            Dim strLattitude As String = station("Lattitude")
            Dim strLongitude As String = station("Longitude")

            Dim strLocationValue As String = String.Format("{0},{1}", strLattitude, strLongitude)

            Dim strStationDesc As String = String.Empty

            If String.IsNullOrEmpty(station("Distance")) = False Then
              strStationDesc = String.Format("{0} {1} away", station("Distance"), station("Unit").ToLower)
              strStationName = String.Format("{0} [{1}]", station("StationName"), strStationDesc)
            End If

            selSearchStationId.AddItem(strStationName, strLocationValue, strLocationValue = SearchStationId)
            If strLocationValue = SearchStationId Then
              stb.AppendFormat("<input type='hidden' name='txtStationName' value='{0}'>{1}", strStationName, vbCrLf)
            End If

          Next

          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Select Station Id</td>")
          stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;{1}</td>{2}", selSearchStationId.Build(), jbSearchStationId.Build(), vbCrLf)
          stb.AppendLine(" </tr>")

        End If

      End If

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStations")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Devices Tab
  ''' </summary>
  ''' <param name="strDeviceType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabDevices(Optional ByVal strDeviceType As String = "Temperature", Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmDevices", "frmDevices", "Post"))

      '
      ' WeatherBug Device Types
      '
      Dim selWeatherTypes As New clsJQuery.jqDropList("selWeatherTypes", Me.PageName, False)
      selWeatherTypes.id = "selWeatherTypes"
      selWeatherTypes.toolTip = "Select the WeatherBug Device Type."

      Dim WeatherTypes As String() = {"Temperature", "Humidity", "Wind", "Rain", "Pressure", "Visibility", "Weather", "Forecast", "Alerts"}
      For Each WeatherType In WeatherTypes
        Dim value As String = WeatherType
        Dim desc As String = WeatherType
        selWeatherTypes.AddItem(desc, value, WeatherType = strDeviceType)
      Next

      stb.AppendLine(" <div>")
      stb.AppendFormat("<b>{0}:</b>&nbsp;{1}{2}", "WeatherBug Device Type", selWeatherTypes.Build, vbCrLf)
      stb.AppendLine(" </div>")

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      '
      ' WeatherBug Devices
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='6'>WeatherBug Devices</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecolumn'>Station #</td>")
      stb.AppendLine("  <td class='tablecolumn'>Name</td>")
      stb.AppendLine("  <td class='tablecolumn'>Type</td>")
      stb.AppendLine("  <td class='tablecolumn'>Value</td>")
      stb.AppendLine("  <td class='tablecolumn'>Last Change</td>")
      stb.AppendLine("  <td class='tablecolumn'>HomeSeer Device <input type=""checkbox"" title=""Check/Uncheck All""onclick=""javascript:toggleAll('chkAddDevice', this.checked)""/></td>")
      stb.AppendLine(" </tr>")

      Dim Keys() As String = hspi_plugin.GetWeatherKeys(strDeviceType)
      For Each strStationNumber As String In Stations.Keys
        Dim strStationName As String = Trim(GetSetting(strStationNumber, "StationName", ""))
        If strStationName.Length > 0 Then

          For Each strKey As String In Keys
            '
            ' Get the weatherbug data
            '
            Dim objWeatherData As Hashtable = Stations(strStationNumber)(strKey)

            Dim strChkInputName As String = "chkAddDevice"
            Dim strChkDisabled As String = ""
            Dim strChecked As String = ""
            Dim strHSDevice As String = ""

            If objWeatherData("Name") = "Unknown" Then
              strChkDisabled = "disabled"
              strChecked = ""
            ElseIf objWeatherData("DevCode") = "" Then
              strChkDisabled = ""
              strChecked = ""
            Else
              strChkDisabled = "disabled"
              strChecked = "checked"
            End If

            Dim strChkInputValue As String = String.Format("{0}:{1}", strStationNumber, strKey)
            strHSDevice = String.Format("<input type='{0}' name='{1}' value='{2}' {3} {4}>", "checkbox", strChkInputName, strChkInputValue, strChecked, strChkDisabled)

            stb.AppendLine(" <tr>")
            stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", strStationNumber)
            stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", objWeatherData("Name"))
            stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", objWeatherData("Type"))
            stb.AppendFormat("<td class='{0}' align='right'>{1}</td>", "tablecell", objWeatherData("Value"))
            stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", objWeatherData("LastChange").ToString)
            stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", strHSDevice)
            stb.AppendLine(" </tr>")

          Next
        End If
      Next

      stb.AppendLine("</table")

      Dim jqButton1 As New clsJQuery.jqButton("btnAddDevices", "Save", Me.PageName, True)
      stb.AppendLine("<div>")
      stb.AppendLine(jqButton1.Build())
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divDevices", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabDevices")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Live Weather Tab
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabLiveWeather(Optional ByVal strStationNumber As String = "Station1", Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      '
      ' WeatherBug Stations
      '
      Dim selLiveWeaatherStation As New clsJQuery.jqDropList("selLiveWeaatherStation", Me.PageName, False)
      selLiveWeaatherStation.id = "selLiveWeaatherStation"
      selLiveWeaatherStation.toolTip = "Select the WeatherBug Station."

      For i As Integer = 1 To 5
        Dim value As String = String.Format("Station{0}", i.ToString)
        Dim desc As String = String.Format("Station {0}", i.ToString)
        selLiveWeaatherStation.AddItem(desc, value, value = strStationNumber)
      Next

      stb.AppendLine(" <div>")
      stb.AppendFormat("<b>{0}:</b>&nbsp;{1}{2}", "WeatherBug Station", selLiveWeaatherStation.Build, vbCrLf)
      stb.AppendLine(" </div>")

      Dim Station As Hashtable = hspi_plugin.GetStation(strStationNumber)
      If Station.Count = 0 Then
        stb.AppendLine("No WeatherBug station devices found for selected station.")
      Else

        Dim strStationName As String = Trim(GetSetting(strStationNumber, "StationName", ""))
        Dim strConditionIcon As String = Station("current-condition")("Image")
        Dim strConditionImage As String = "/images/hspi_ultraweatherbug3/sunny.png"

        stb.AppendLine("<table border='0' cellspacing='10' width='100%'>")
        stb.AppendLine(" <tr>")
        stb.AppendFormat("  <td class='tableheader' colspan='3'>{0}</td>", String.Format("Live Conditions {0}", Station("ob-date")("Value")))
        stb.AppendLine(" </tr>")
        stb.AppendLine(" <tr>")
        stb.AppendFormat("  <td class='tablecolumn' colspan='3'>{0}</td>", strStationName)
        stb.AppendLine(" </tr>")
        stb.AppendLine(" <tr>")

        stb.AppendLine("  <td class='tablecell' style='vertical-align:top; width:33.3%;'>")
        stb.AppendLine("   <table border='0' style='width:100%'>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td class='tablecell' style='text-align:center;' colspan='2'>Temperature</td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td style='text-align:center; vertical-align:top; font-size: 20px;' colspan='2'>")
        stb.AppendFormat("     <div><img width='32' height='32' align='absmiddle' src='/images/hspi_ultraweatherbug3/temperature.png'/>&nbsp;{0}</div>", String.Format("{0} {1}", Station("temp")("Value"), Station("temp")("Units")))
        stb.AppendLine("     </td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td style='text-align:center; vertical-align:top; font-size: 16px' colspan='2'>Feels Like {0}</td>", Station("feels-like")("String"))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td colspan='2'>&nbsp;</td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("temp-rate")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("temp-rate")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("heat-index")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("heat-index")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("dew-point")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("dew-point")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("dew-point-rate")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("dew-point-rate")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("humidity")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("humidity")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("humidity-rate")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("humidity-rate")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("   </table>")
        stb.AppendLine("  </td>")

        stb.AppendLine("  <td class='tablecell' style='vertical-align:top; width:33.3%;'>")
        stb.AppendLine("   <table border='0' style='width:100%'>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td class='tablecell' style='text-align:center;' colspan='2'>Winds</td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td style='text-align:center; vertical-align:top; font-size: 20px;' colspan='2'>")
        stb.AppendFormat("     <div>{0}&nbsp;{1}</div>", Station("wind-direction")("Image"), String.Format("{0}", Station("wind-direction")("String")))
        stb.AppendLine("     </td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td style='text-align:center; vertical-align:top; font-size: 16px' colspan='2'>{0}</td>", Station("wind-speed")("String"))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td colspan='2'>&nbsp;</td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("wind-direction-avg")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("wind-direction-avg")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("wind-speed-avg")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("wind-speed-avg")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("gust-direction")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("gust-direction")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("gust-speed")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("gust-speed")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("   </table>")
        stb.AppendLine("  </td>")

        stb.AppendLine("  <td class='tablecell' style='vertical-align:top; width:33.3%;'>")
        stb.AppendLine("   <table border='0' style='width:100%'>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td class='tablecell' style='text-align:center;' colspan='2'>Conditions</td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td style='text-align:center; vertical-align:top; font-size: 20px;' colspan='2'>")
        stb.AppendFormat("     <div>{0}&nbsp;{1}</div>", strConditionIcon, Station("current-condition")("String"))
        stb.AppendLine("     </td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td style='text-align:center; vertical-align:top; font-size: 16px' colspan='2'>High {0} Low {1}</td>", FormatDeviceValue(Station("todays-temperature-day"), False), FormatDeviceValue(Station("todays-temperature-night"), False))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendLine("     <td colspan='2'>&nbsp;</td>")
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("rain-today")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("rain-today")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("rain-rate")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("rain-rate")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("rain-month")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("rain-month")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("rain-year")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("rain-year")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("pressure")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("pressure")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("    <tr>")
        stb.AppendFormat("     <td>{0}</td>", Station("pressure-rate")("Name"))
        stb.AppendFormat("     <td nowrap>{0}</td>", FormatDeviceValue(Station("pressure-rate")))
        stb.AppendLine("    </tr>")
        stb.AppendLine("   </table>")
        stb.AppendLine("  </td>")

        stb.AppendLine(" </tr>")
        stb.AppendLine("</table")

        Dim jqButton1 As New clsJQuery.jqButton("btnRefreshLiveWeather", "Refresh", Me.PageName, True)
        stb.AppendLine("<div>")
        stb.AppendLine(jqButton1.Build())
        stb.AppendLine("</div>")

      End If

      'stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divLiveWeather", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabLiveWeather")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Live Weather Tab
  ''' </summary>
  ''' <param name="strStationNumber"></param>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabForecast(Optional ByVal strStationNumber As String = "Station1", Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      '
      ' WeatherBug Stations
      '
      Dim selForecastStations As New clsJQuery.jqDropList("selForecastStations", Me.PageName, False)
      selForecastStations.id = "selForecastStations"
      selForecastStations.toolTip = "Select the WeatherBug Station."

      For i As Integer = 1 To 5
        Dim value As String = String.Format("Station{0}", i.ToString)
        Dim desc As String = String.Format("Station {0}", i.ToString)
        selForecastStations.AddItem(desc, value, value = strStationNumber)
      Next

      stb.AppendLine(" <div>")
      stb.AppendFormat("<b>{0}:</b>&nbsp;{1}{2}", "WeatherBug Station", selForecastStations.Build, vbCrLf)
      stb.AppendLine(" </div>")

      Dim Forecasts As Hashtable = hspi_plugin.GetForecast(strStationNumber)
      If Forecasts.Count = 0 Then
        stb.AppendLine("No WeatherBug forecast found for selected station.")
      Else

        Dim DaysOfWeek As ArrayList = hspi_plugin.GetDaysOfWeek()
        Dim strStationName As String = Trim(GetSetting(strStationNumber, "StationName", ""))

        stb.AppendLine("<table border='0' cellspacing='0' width='100%'>")
        stb.AppendLine(" <tr>")
        stb.AppendFormat("  <td class='tableheader' colspan='14'>{0} - {1}</td>", strStationNumber, strStationName)
        stb.AppendLine(" </tr>")

        stb.AppendLine(" <tr>")

        Dim strTemperatureIcon As String = "<img width='16' height='16' src='/images/hspi_ultraweatherbug3/temperature.png'>"
        Dim strHumidityIcon As String = "<img width='16' height='16' src='/images/hspi_ultraweatherbug3/humidity.png'>"
        Dim strWindsIcon As String = "<img width='16' height='16' src='/images/hspi_ultraweatherbug3/wind.png'>"


        For i As Integer = 1 To 8
          Dim Forecast As Hashtable = Forecasts(i)

          Dim txtForecastTitle As String = Forecast("title")
          Dim txtPrediction As String = Forecast("summaryDescription")
          Dim txtDetailedDescription As String = Forecast("detailedDescription")

          Dim txtHighLow As String = Forecast("temperature")
          Dim txtDewPoint As String = String.Format("{0} {1}", "Dew Point", Forecast("dewPoint"))
          Dim txtHumidity As String = String.Format("{0} {1}", "Humidity", Forecast("relativeHumidity"))
          Dim txtWindDir As String = String.Format("{0} {1}", "Winds", Forecast("windDirectionDegrees"))
          Dim txtWindSpeed As String = String.Format("{0}", Forecast("windSpeed"))

          stb.AppendLine("  <td class='tablecell' style='vertical-align:top; width:12.5%;'>")
          stb.AppendLine("   <table style='border-collapse:collapse'>")
          stb.AppendLine("    <tr>")
          stb.AppendFormat("   <td class='tablecell' style='text-align:center;' colspan='2'>{0}</td>", txtForecastTitle)
          stb.AppendLine("    </tr>")

          stb.AppendLine("    <tr>")
          stb.AppendLine("     <td style='height:80px; text-align:center; vertical-align:top;'>")
          stb.AppendFormat("    <div>{0}</div>", Forecast("iconImage"))
          stb.AppendFormat("    <div>{0}</div>", txtPrediction)
          stb.AppendLine("     </td>")
          stb.AppendLine("    </tr>")

          stb.AppendLine("    <tr>")
          stb.AppendFormat("   <td style='height:140px; text-align:center; vertical-align:top'>")
          stb.AppendFormat("    <div style='font-size: 30px; height:60px;'>{0}</div>", txtHighLow)
          stb.AppendFormat("    <div style='font-size: 12px;'>{0}</div>", txtDewPoint)
          stb.AppendFormat("    <div style='font-size: 12px;'>{0}</div>", txtHumidity)
          stb.AppendFormat("    <div style='font-size: 12px;'>{0}</div>", txtWindDir)
          stb.AppendFormat("    <div style='font-size: 12px;'>{0}</div>", txtWindSpeed)
          stb.AppendLine("     </td>")
          stb.AppendLine("    </tr>")

          stb.AppendLine("    <tr>")
          stb.AppendFormat("   <td style='text-align:left; vertical-align:top'>")
          stb.AppendFormat("    <div>{0}</div>", txtDetailedDescription)
          stb.AppendLine("     </td>")
          stb.AppendLine("    </tr>")

          stb.AppendLine("  </table")
          stb.AppendLine(" </td>")

        Next

        stb.AppendLine(" </tr>")
        stb.AppendLine("</table")

      End If

      If Rebuilding Then Me.divToUpdate.Add("divForecast", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabForecast")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Alerts Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabAlerts(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmAlerts", "frmAlerts", "Post"))

      '
      ' Get the WeatherBug Alerts for each station
      '
      Dim Alerts As SortedList = hspi_plugin.Alerts
      If Alerts.Count = 0 Then
        stb.AppendLine("There are no weather alerts at this time for the selected station.")
      Else

        stb.AppendLine("<table cellspacing='0' width='100%'>")

        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class='tableheader' colspan='6'>WeatherBug Alerts</td>")
        stb.AppendLine(" </tr>")

        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class='tablecolumn'>Alert #</td>")
        stb.AppendLine("  <td class='tablecolumn'>Alert Type</td>")
        stb.AppendLine("  <td class='tablecolumn'>Date Posted</td>")
        stb.AppendLine("  <td class='tablecolumn'>Date Expires</td>")
        stb.AppendLine("  <td class='tablecolumn'>Title</td>")
        stb.AppendLine("  <td class='tablecolumn'>Alert Message Summary</td>")
        stb.AppendLine(" </tr>")

        For Each strAlertId As String In Alerts.Keys
          stb.AppendLine(" <tr>")
          stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", strAlertId)
          stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", Alerts(strAlertId)("type"))
          stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", Alerts(strAlertId)("posted-date"))
          stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", Alerts(strAlertId)("expires-date"))
          stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", Alerts(strAlertId)("title"))
          stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", Alerts(strAlertId)("msg-summary"))
          stb.AppendLine(" </tr>")
        Next

        stb.AppendLine("</table")
      End If


      If Rebuilding Then Me.divToUpdate.Add("divAlerts", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabAlerts")
      Return "error - " & Err.Description
    End Try

  End Function

  Function CheckPulseAPIKey() As String

    Try

      If WeatherBugAPI.ValidAPIKey = False Then

        Dim stb As New StringBuilder

        stb.AppendLine("<h3>UltraWeatherBug3 HSPI Setup</h3>")
        stb.AppendLine("<p>Before you begin to configure the UltraWeatherBug3 plug-in, you’ll need to register for a WeatherBug - Earth Networks API key.</p>")
        stb.AppendLine("<p>How to Purchase Your WeatherBug Pulse API Key:</p>")
        stb.AppendLine("<ol>")
        stb.AppendLine("<li>Sign up for a WeatherBug User account at https://developer.earthnetworks.com/pricing/ </li>")
        stb.AppendLine("<li>Check your inbox and follow the instructions WeatherBug sent you by email.</li>")
        stb.AppendLine("<li>Click the ""Subscriptions"" option in the toolbar, then click on Subscribe to our Pulse API link.</li>")
        stb.AppendLine("<li>Select Basic $20 per month plan, then click subscribe now and follow the onscreen instructions to complete payment.</li>")
        stb.AppendLine("<li>Once you have successfully subscribed to a plan, click the Manage API option in the toolbar.</li>")
        stb.AppendLine("<li>From the UltraWeatherBug > Options > WeatherBug API Key section, copy and paste the Primary Key and Secondary Key into the text boxes provided, then click Save Options.</li>")
        stb.AppendLine("</ol>")

        Return stb.ToString

      End If

    Catch pEx As Exception

    End Try

    Return ""

  End Function

  ''' <summary>
  ''' Formats the Device Value for Display
  ''' </summary>
  ''' <param name="device"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Function FormatDeviceValue(device As Hashtable, Optional displayImage As Boolean = True) As String

    Try

      If device.Contains("Value") Then
        If displayImage = True Then
          Return String.Format("{0}&nbsp;{1}{2}", device("Icon"), device("Value"), device("Units"))
        Else
          Return String.Format("{0}{1}", device("Value"), device("Units"))
        End If
      Else
        Return "--"
      End If

    Catch pEx As Exception
      Return device("Value")
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabTemplates"
          BuildTabTemplates(True)

        Case "tabStations"
          Dim strPulseAPIKey As String = CheckPulseAPIKey()
          If strPulseAPIKey.Length > 0 Then
            Me.divToUpdate.Add("divStations", strPulseAPIKey)
          Else
            BuildTabStations(True)
          End If

        Case "tabDevices"
          Dim strPulseAPIKey As String = CheckPulseAPIKey()
          If strPulseAPIKey.Length > 0 Then
            Me.divToUpdate.Add("divStations", strPulseAPIKey)
          Else
            BuildTabDevices("Temperature", True)
          End If

        Case "tabLiveWeather"
          Dim strPulseAPIKey As String = CheckPulseAPIKey()
          If strPulseAPIKey.Length > 0 Then
            Me.divToUpdate.Add("divStations", strPulseAPIKey)
          Else
            BuildTabLiveWeather("Station1", True)
          End If

        Case "tabForecast"
          Dim strPulseAPIKey As String = CheckPulseAPIKey()
          If strPulseAPIKey.Length > 0 Then
            Me.divToUpdate.Add("divStations", strPulseAPIKey)
          Else
            BuildTabForecast("Station1", True)
          End If

        Case "tabAlerts"
          Dim strPulseAPIKey As String = CheckPulseAPIKey()
          If strPulseAPIKey.Length > 0 Then
            Me.divToUpdate.Add("divStations", strPulseAPIKey)
          Else
            BuildTabAlerts(True)
          End If

        Case "txtAPIKeyPrimary"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("API", "KeyPrimary", strValue)

          PostMessage("The Pulse API Primary Key has been updated.  A restart of HomeSeer may be required.")

        Case "txtAPIKeySecondary"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("API", "KeySecondary", strValue)

          PostMessage("The Pulse API Secondary Key has been updated.  A restart of HomeSeer may be required.")

        Case "selUnitType"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Options", "UnitType", strValue)

          PostMessage("The Unit Type has been updated.")

        Case "selLiveWeatherUpdate"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Options", "LiveWeatherUpdate", strValue)

          PostMessage("The Station Update Frequency has been updated.")

        Case "selForecastUpdate"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Options", "ForecastUpdate", strValue)

          PostMessage("The Forecast Update Frequency has been updated.")

        Case "selAlertsUpdate"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Options", "AlertsUpdate", strValue)

          PostMessage("The Alerts Update Frequency has been updated.")

        Case "selDeviceImage"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Options", "DeviceImage", strValue)

          PostMessage("The HomeSeer Device Option has been updated.")

        Case "txtEmailRcptTo"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailRcptTo", strValue)

          PostMessage("The Email Recipient Option has been updated.")

        Case "txtEmailFrom"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailFrom", strValue)

          PostMessage("The Email Sender Option has been updated.")

        Case "txtEmailSubject"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("EmailNotification", "EmailSubject", strValue)

          PostMessage("The Email Subject Option has been updated.")

        Case "btnSaveTemplates"

          Dim strValue As String = postData("txtWeatherCurrent").Trim.Replace(vbCrLf, "~")
          SaveSetting("EmailNotification", "WeatherCurrent", strValue)

          strValue = postData("txtWeatherForecast").Trim.Replace(vbCrLf, "~")
          SaveSetting("EmailNotification", "WeatherForecast", strValue)

          strValue = postData("txtWeatherAlert").Trim.Replace(vbCrLf, "~")
          SaveSetting("EmailNotification", "WeatherAlert", strValue)

          strValue = postData("txtSpeakWeatherTemplate").Trim.Replace(vbCrLf, "~")
          SaveSetting("Speak", "WeatherCurrent", strValue)

          strValue = postData("txtSpeakWeatherForecast").Trim.Replace(vbCrLf, "~")
          SaveSetting("Speak", "WeatherForecast", strValue)

          strValue = postData("txtSpeakWeatherAlert").Trim.Replace(vbCrLf, "~")
          SaveSetting("Speak", "WeatherAlert", strValue)

          PostMessage("The Weather Templates have been updated.")

        Case "chkWeatherCurrent"
          SaveSetting("EmailNotification", "WeatherCurrent", WEATHER_CURRENT_TEMPLATE)
          Me.divToUpdate.Add("divTemplates", BuildTabTemplates())

        Case "chkWeatherForecast"
          SaveSetting("EmailNotification", "WeatherForecast", WEATHER_FORECAST_TEMPLATE)
          Me.divToUpdate.Add("divTemplates", BuildTabTemplates())

        Case "chkWeatherAlert"
          SaveSetting("EmailNotification", "WeatherAlert", WEATHER_ALERT_TEMPLATE)
          Me.divToUpdate.Add("divTemplates", BuildTabTemplates())

        Case "chkSpeakWeatherCurrent"
          SaveSetting("Speak", "WeatherCurrent", SPEAK_WEATHER_CURRENT)
          Me.divToUpdate.Add("divTemplates", BuildTabTemplates())

        Case "chkSpeakWeatherForecast"
          SaveSetting("Speak", "WeatherForecast", SPEAK_WEATHER_FORECAST)
          Me.divToUpdate.Add("divTemplates", BuildTabTemplates())

        Case "chkSpeakWeatherAlert"
          SaveSetting("Speak", "WeatherAlert", SPEAK_WEATHER_ALERT)
          Me.divToUpdate.Add("divTemplates", BuildTabTemplates())

        Case "btnUpdateStation1"
          Me.divToUpdate.Add("divStations", BuildTabStationUpdate("Station1", "btnUpdateStation", "", "", ""))

        Case "btnUpdateStation2"
          Me.divToUpdate.Add("divStations", BuildTabStationUpdate("Station2", "btnUpdateStation", "", "", ""))

        Case "btnUpdateStation3"
          Me.divToUpdate.Add("divStations", BuildTabStationUpdate("Station3", "btnUpdateStation", "", "", ""))

        Case "btnUpdateStation4"
          Me.divToUpdate.Add("divStations", BuildTabStationUpdate("Station4", "btnUpdateStation", "", "", ""))

        Case "btnUpdateStation5"
          Me.divToUpdate.Add("divStations", BuildTabStationUpdate("Station5", "btnUpdateStation", "", "", ""))

        Case "btnSearchLocation"
          Dim strStationNumber As String = postData("txtStationNumber")
          Dim strSearchLocation As String = postData("txtSearchLocation")

          Me.divToUpdate.Add("divStations", BuildTabStationUpdate(postData("txtStationNumber"), postData("id"), strSearchLocation, "", ""))

        Case "btnSearchCity"
          Dim strStationNumber As String = postData("txtStationNumber")
          Dim strSearchLocation As String = postData("txtSearchLocation")
          Dim strSearchCity As String = postData("selSearchCity")

          Me.divToUpdate.Add("divStations", BuildTabStationUpdate(postData("txtStationNumber"), postData("id"), strSearchLocation, strSearchCity, ""))

        Case "btnSearchStationId"
          Dim strStationNumber As String = postData("txtStationNumber")
          Dim strSearchLocation As String = postData("txtSearchLocation")
          Dim strSearchCity As String = postData("selSearchCity")
          Dim strStationId As String = postData("selSearchStationId")

          Me.divToUpdate.Add("divStations", BuildTabStationUpdate(postData("txtStationNumber"), postData("id"), strSearchLocation, strSearchCity, strStationId))

        Case "btnSetStationId"
          Dim strStationNumber As String = postData("txtStationNumber")
          Dim strSearchLocation As String = postData("txtSearchLocation")
          Dim strCityId As String = postData("selSearchCity")
          Dim strCityName As String = postData("txtCityName")
          Dim strStationId As String = postData("selSearchStationId")
          Dim strStationName As String = postData("txtStationName")

          SaveSetting(strStationNumber, "CityId", strCityId)
          SaveSetting(strStationNumber, "CityName", strCityName)

          SaveSetting(strStationNumber, "StationId", strStationId)
          SaveSetting(strStationNumber, "StationName", strStationName)

          Me.divToUpdate.Add("divStations", BuildTabStations())

        Case "btnClearStation2"
          SaveSetting("Station2", "CityId", "")
          SaveSetting("Station2", "CityName", "")

          SaveSetting("Station2", "StationId", "")
          SaveSetting("Station2", "StationName", "")

          Me.divToUpdate.Add("divStations", BuildTabStations())

        Case "btnClearStation3"
          SaveSetting("Station3", "CityId", "")
          SaveSetting("Station3", "CityName", "")

          SaveSetting("Station3", "StationId", "")
          SaveSetting("Station3", "StationName", "")

          Me.divToUpdate.Add("divStations", BuildTabStations())

        Case "btnClearStation4"
          SaveSetting("Station4", "CityId", "")
          SaveSetting("Station4", "CityName", "")

          SaveSetting("Station4", "StationId", "")
          SaveSetting("Station4", "StationName", "")

          Me.divToUpdate.Add("divStations", BuildTabStations())

        Case "btnClearStation5"
          SaveSetting("Station5", "CityId", "")
          SaveSetting("Station5", "CityName", "")

          SaveSetting("Station5", "StationId", "")
          SaveSetting("Station5", "StationName", "")

          Me.divToUpdate.Add("divStations", BuildTabStations())

        Case "selWeatherTypes"
          Dim strValue As String = postData(postData("id"))
          BuildTabDevices(strValue, True)

        Case "btnAddDevices"
          Dim strValues As String = postData("chkAddDevice")

          '
          ' Add each selected devices
          '      
          If Not (strValues Is Nothing) Then
            For Each Item As Object In strValues.Split(",")
              Dim strResults As String = hspi_devices.CreateWeatherBugDevice(Item, False)

              If strResults <> String.Empty Then
                PostMessage("Failed to process selected WeatherBug Devices due to an error: " & strResults)
              End If
            Next

          End If

          PostMessage("<p>Device configuration saved.</p>")

          BuildTabDevices(postData("selWeatherTypes"), True)

        Case "selLiveWeaatherStation", "btnRefreshLiveWeather"
          Dim strValue As String = postData("selLiveWeaatherStation")
          BuildTabLiveWeather(strValue, True)

        Case "selForecastStations"
          Dim strValue As String = postData("selForecastStations")
          BuildTabForecast(strValue, True)

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class

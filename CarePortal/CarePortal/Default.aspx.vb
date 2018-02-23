Imports OboutInc.EasyMenu_Pro
Imports System.Web.Configuration
Imports System.Xml

<CLSCompliant(False)>
Partial Public Class _Default
  Inherits System.Web.UI.Page

  Private mvCareControls As List(Of CareWebControl)

  Private Sub Page_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit
    Dim vPageNumber As Integer
    Dim vLoginPageNumber As Integer
    Dim vUpdateDetailsPageNumber As Integer
    Dim vSiteLogo As String = ""

    Try
      If User.Identity.IsAuthenticated AndAlso
         TypeOf (User.Identity) Is System.Security.Principal.WindowsIdentity AndAlso
         Session("UserLogname") Is Nothing Then
        Try
          Dim vList As New ParameterList
          Dim vLogname As String = CType(User.Identity, System.Security.Principal.WindowsIdentity).Name
          Dim vPos As Integer = vLogname.IndexOf("\")
          If vPos >= 0 Then vLogname = vLogname.Substring(vPos + 1)
          vList("UserName") = vLogname
          vList("Password") = "none"
          vList("AuthenticatedUser") = vLogname
          Dim vReturnList As ParameterList = DataHelper.Login(vList)
          Session("UserLogname") = vReturnList("UserLogname")
          Session("UserContactNumber") = vReturnList("ContactNumber")
          Session("UserAddressNumber") = vReturnList("AddressNumber")
          Session("UserDepartment") = vReturnList("UserDepartment")
          Session("Database") = DataHelper.Database
        Catch vEx As Exception
          ProcessError(vEx)
        End Try
      End If

      Dim vTable As DataTable = GetWebInfo()
      If vTable IsNot Nothing Then
        vPageNumber = IntegerValue(vTable.Rows(0)("WebPageNumber").ToString)
        vLoginPageNumber = IntegerValue(vTable.Rows(0)("LoginPageNumber").ToString)
        vUpdateDetailsPageNumber = IntegerValue(vTable.Rows(0)("UpdateDetailsPageNumber").ToString)
        Dim vHeaderHtml As String = vTable.Rows(0)("HeaderHtml").ToString
        If vHeaderHtml.Length > 0 Then SiteHeader.InnerHtml = vHeaderHtml
        Dim vFooterHtml As String = vTable.Rows(0)("FooterHtml").ToString
        If vFooterHtml.Length > 0 Then SiteFooter.InnerHtml = vFooterHtml
        Dim vLeftPanelHtml As String = vTable.Rows(0)("LeftPanelHtml").ToString
        If vLeftPanelHtml.Length > 0 Then SiteLeftPanel.InnerHtml = vLeftPanelHtml
        Dim vRightPanelHtml As String = vTable.Rows(0)("RightPanelHtml").ToString
        If vRightPanelHtml.Length > 0 Then SiteRightPanel.InnerHtml = vRightPanelHtml
      End If
      If Request.QueryString("pn") IsNot Nothing Then
        vPageNumber = IntegerValue(Request.QueryString("pn"))
      Else
        Session("LastPageNumber") = vPageNumber
      End If

      If vPageNumber > 0 Then
        Dim vPageTitle As String = ""
        Dim vPageMenu As Integer
        Dim vAccessViewName As String = String.Empty
        Dim vPageTable As DataTable = GetPageInfo(vPageNumber)
        If vPageTable IsNot Nothing Then
          If vPageTable.Columns.Contains("PagePublished") AndAlso InWebPageDesigner() = False Then
            If vPageTable.Rows(0).Item("PagePublished").ToString = "N" Then Throw New Exception(String.Format("The requested page {0} has not been published and cannot be displayed", vPageNumber))
          End If
          If vPageTable.Columns.Contains("LoginTypeRequired") Then
            Dim vLoginTypeRequired As String = vPageTable.Rows(0).Item("LoginTypeRequired").ToString
            If vPageTable.Columns.Contains("AccessViewName") Then vAccessViewName = vPageTable.Rows(0).Item("AccessViewName").ToString
            Select Case vLoginTypeRequired
              Case "R"      'Registered User
                CheckAuthentication(False, vLoginPageNumber, vAccessViewName)
              Case "U"      'CARE User
                CheckAuthentication(True, vLoginPageNumber)
            End Select
          End If
          If vPageTable.Columns.Contains("SuppressSiteHeader") Then
            SiteHeader.Visible = vPageTable.Rows(0).Item("SuppressSiteHeader").ToString = "N"
            SiteFooter.Visible = vPageTable.Rows(0).Item("SuppressSiteFooter").ToString = "N"
          End If
          If vPageTable.Columns.Contains("SuppressSiteLeftPanel") Then
            SiteLeftPanel.Visible = vPageTable.Rows(0).Item("SuppressSiteLeftPanel").ToString = "N"
          End If
          If vPageTable.Columns.Contains("SuppressSiteRightPanel") Then
            SiteRightPanel.Visible = vPageTable.Rows(0).Item("SuppressSiteRightPanel").ToString = "N"
          End If
          vPageTitle = vPageTable.Rows(0).Item("WebPageTitle").ToString
          vPageMenu = IntegerValue(vPageTable.Rows(0).Item("WebMenuNumber").ToString)
          Session("FriendlyUrl") = vPageTable.Rows(0).Item("FriendlyUrl").ToString
        End If
        Me.Title = vPageTitle
        If vPageMenu > 0 Then GetMenus(vPageMenu)
        GetPageContent(vPageNumber, vLoginPageNumber, vUpdateDetailsPageNumber)
      End If
      'to stop users navigating to another page before updating their details
      If Session("UpdateDetailsURL") IsNot Nothing AndAlso vUpdateDetailsPageNumber <> vPageNumber Then
        ProcessRedirect(Session("UpdateDetailsURL").ToString)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Debug.Print("Got Page Load")
    'BR15415 The following line fixes the form action after a URL ReWrite so that it correctly redirects after a submit
    nfpform.Action = Request.RawUrl
    Dim vHeaderText As String = Utilities.GetCustomPageElement("head")
    'If <title> tag is not specified in the custom head.htm file then add attriute <title> to the 
    'head.htm file and add the page title value as the title specified in WPD while configuring
    ' the page. It the <title> tag is specified in the custom head.htm and there is no value specified
    ' then add the page tile value as the title specified in WPD while configuring the page.
    If Not String.IsNullOrEmpty(vHeaderText) Then
      'Convert XML attribute for <Title> to lowercase as user can input tags in any format
      If CBool(InStr(vHeaderText, "<title>", CompareMethod.Text)) Then
        Dim vString1 As String = System.Text.RegularExpressions.Regex.Replace(vHeaderText, "<title>", "<title>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim vString2 As String = System.Text.RegularExpressions.Regex.Replace(vString1, "</title>", "</title>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        vHeaderText = vString2
      End If
      If CBool(InStr(vHeaderText, "<head>", CompareMethod.Text)) Then
        vHeaderText = System.Text.RegularExpressions.Regex.Replace(vHeaderText, "<head>", "<head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        vHeaderText = System.Text.RegularExpressions.Regex.Replace(vHeaderText, "</head>", "</head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim vHeaderIndex As Integer = vHeaderText.IndexOf("<head>")
        If vHeaderIndex < vHeaderText.Length Then vHeaderText = vHeaderText.Substring(vHeaderIndex)
      End If

      Dim vXmlSearch As New XmlDocument()
      Try
        vXmlSearch.LoadXml(vHeaderText)
      Catch vEx As Exception
        ProcessError(vEx)
      End Try

      Select Case vXmlSearch.GetElementsByTagName("title").Count
        Case 0
          Dim vXmlNode As XmlNode = vXmlSearch.DocumentElement
          Dim vXmlElement As XmlElement = vXmlSearch.CreateElement("title")
          vXmlElement.InnerText = Page.Header.Title
          vXmlNode.AppendChild(vXmlElement)
          vHeaderText = vXmlSearch.OuterXml
        Case Is > 0
          If vXmlSearch.DocumentElement("title").InnerText.Length = 0 Then
            vXmlSearch.DocumentElement("title").InnerText = Page.Header.Title
            vHeaderText = vXmlSearch.OuterXml
          End If
        Case Else
          'Do nothing as properheade value is declared in the head.htm file
      End Select
      If Not String.IsNullOrEmpty(vHeaderText) Then Page.Header.InnerHtml = vHeaderText
    End If

    BodyStart.Text = Utilities.GetCustomPageElement("bodystart")
    BodyEnd.Text = Utilities.GetCustomPageElement("bodyend")
  End Sub

  Public Sub CheckAuthentication(ByVal pCareLogin As Boolean, ByVal pLoginPageNumber As Integer, Optional ByVal pAccessViewName As String = "")
    If Request.QueryString("cwpd") <> "Y" Then
      Dim vDepartment As String = ""
      If HttpContext.Current.User.Identity.IsAuthenticated Then
        If TypeOf (User.Identity) Is System.Security.Principal.WindowsIdentity Then
          vDepartment = Session("UserDepartment").ToString
        Else
          Dim vIdentity As FormsIdentity = CType(User.Identity, FormsIdentity)
          If vIdentity.Ticket.UserData.Length > 0 Then
            Dim vItems As String() = vIdentity.Ticket.UserData.Split("|"c)
            If vItems.Length > 3 Then vDepartment = vItems(3)
            If pAccessViewName.Length > 0 Then 'Access to the page is restricted
              If vItems.Length > 4 Then 'Check if viewname exists in userdata
                'Split again to get a list of views that the user belongs to
                Dim vViews As String() = vItems(4).Split(","c)
                If Array.IndexOf(vViews, pAccessViewName) = -1 Then Throw New PortalAccessException()
              End If
            End If
          End If
        End If
      End If
      If HttpContext.Current.User.Identity.IsAuthenticated = False OrElse (pCareLogin AndAlso vDepartment.Length = 0) Then
        If pLoginPageNumber > 0 Then
          ProcessRedirect(String.Format("default.aspx?pn={0}&ReturnURL={1}&Type={2}", pLoginPageNumber, Server.UrlEncode(Request.Url.ToString), pCareLogin))
        Else
          Throw New CareException("Web Page requires an Authenticated User but no login page has been defined")
        End If
      End If
    End If
  End Sub

  Public Sub GetMenus(ByVal pMenuNumber As Integer)
    Dim vList As New ParameterList(HttpContext.Current)
    vList("WebMenuNumber") = pMenuNumber
    Dim vTable As DataTable = DataHelper.GetWebMenus(vList)
    If vTable IsNot Nothing AndAlso vTable.Rows.Count > 0 Then
      Dim vMenuStyle As String = vTable.Rows(0).Item("WebMenuStyle").ToString
      If vMenuStyle.Length > 0 Then
        vMenuStyle = "MenuStyles/" & vMenuStyle
      Else
        vMenuStyle = "MenuStyles/horizontal2"
      End If
      Dim vEM As New EasyMenu
      vEM.ID = "MainMenu"
      vEM.StyleFolder = vMenuStyle                    'set the style for this menu
      vEM.EventsScriptPath = "MenuStyles/Script"
      vEM.Width = "400"
      vEM.ShowEvent = MenuShowEvent.Always            'show event is always so the menu is always visible - this menu doesn't require any AttachTo or Align properties set
      vEM.Position = MenuPosition.Horizontal          'display the menu horizontally
      phcMenu.Controls.Add(vEM)

      'the parent menu looks different so we need to set different
      'CSS classes names for its items and the menu itself
      'css classes names for the menu and the item container
      vEM.CSSMenu = "ParentMenu"
      vEM.CSSMenuItemContainer = "ParentItemContainer"

      'css classes names for MenuItems
      Dim MenuItemCssClasses As CSSClasses = vEM.CSSClassesCollection(vEM.CSSClassesCollection.Add(New CSSClasses(GetType(MenuItem))))
      MenuItemCssClasses.ComponentSubMenuCellOver = "ParentItemSubMenuCellOver"
      MenuItemCssClasses.ComponentContentCell = "ParentItemContentCell"
      MenuItemCssClasses.Component = "ParentItem"
      MenuItemCssClasses.ComponentSubMenuCell = "ParentItemSubMenuCell"
      MenuItemCssClasses.ComponentIconCellOver = "ParentItemIconCellOver"
      MenuItemCssClasses.ComponentIconCell = "ParentItemIconCell"
      MenuItemCssClasses.ComponentOver = "ParentItemOver"
      MenuItemCssClasses.ComponentContentCellOver = "ParentItemContentCellOver"
      'add the classes names to the collection
      vEM.CSSClassesCollection.Add(MenuItemCssClasses)

      'css classes names for MenuSeparators
      Dim MenuSeparatorCssClasses As CSSClasses = vEM.CSSClassesCollection(vEM.CSSClassesCollection.Add(New CSSClasses(GetType(MenuSeparator))))
      MenuSeparatorCssClasses.ComponentSubMenuCellOver = "ParentSeparatorSubMenuCellOver"
      MenuSeparatorCssClasses.ComponentContentCell = "ParentSeparatorContentCell"
      MenuSeparatorCssClasses.Component = "ParentSeparator"
      MenuSeparatorCssClasses.ComponentSubMenuCell = "ParentSeparatorSubMenuCell"
      MenuSeparatorCssClasses.ComponentIconCellOver = "ParentSeparatorIconCellOver"
      MenuSeparatorCssClasses.ComponentIconCell = "ParentSeparatorIconCell"
      MenuSeparatorCssClasses.ComponentOver = "ParentSeparatorOver"
      MenuSeparatorCssClasses.ComponentContentCellOver = "ParentSeparatorContentCellOver"
      'add the classes names to the collection
      vEM.CSSClassesCollection.Add(MenuSeparatorCssClasses)

      Dim vURL As String
      Dim vText As String
      Dim vParentID As String
      Dim vID As String
      Dim vWebPage As Integer
      Dim vSM As EasyMenu = Nothing

      For Each vRow As DataRow In vTable.Rows
        vWebPage = IntegerValue(vRow("WebPageNumber").ToString)
        If vWebPage > 0 Then
          vURL = String.Format("Default.aspx?pn={0}", vWebPage)
        ElseIf vRow("WebUrl").ToString.Length > 0 Then
          vURL = vRow("WebUrl").ToString
        Else
          vURL = ""
        End If
        vText = vRow("MenuTitle").ToString
        Dim vParentNumber As Integer = CInt(vRow("ParentItemNumber"))
        If vParentNumber Mod 100000 = 0 Then
          'Add the menu item to the top level menu
          vParentID = "M" & vRow("WebMenuItemNumber").ToString
          vEM.AddItem(New MenuItem(vParentID, vText, "", vURL))
        Else
          'Look for the parent menu item to add this item to
          Dim vControl As Control = Me.phcMenu.FindControl("SM" & vParentNumber)
          If vControl Is Nothing Then
            vSM = New EasyMenu
            vSM.ID = "SM" & vParentNumber.ToString
            vSM.AttachTo = "M" & vParentNumber.ToString
            vSM.StyleFolder = vMenuStyle
            vSM.EventsScriptPath = "MenuStyles/Script"
            vSM.Width = "150"
            vSM.ShowEvent = MenuShowEvent.MouseOver     'it will show on mouse over
            vSM.Align = MenuAlign.Under                 'and will align under the item to which it is attached
            phcMenu.Controls.Add(vSM)
            vControl = vSM
          End If
          If vControl IsNot Nothing Then
            vID = "pn" & vRow("WebMenuItemNumber").ToString
            DirectCast(vControl, EasyMenu).AddItem(New MenuItem(vID, vText, "", vURL))
          End If
        End If
      Next
    End If
    'If pSelectedID.Length > 0 Then .SelectedIndex = pSelectedID
  End Sub

  Private Function GetWebInfo() As DataTable
    Dim vTable As DataTable = Nothing
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      vList("WebNumber") = WebConfigurationManager.AppSettings("WebNumber").ToString
      vTable = DataHelper.GetWebInfo(vList)
    Catch vException As Exception
      ProcessError(vException)
    End Try
    Return vTable
  End Function

  Private Function GetPageInfo(ByVal pPageNumber As Integer) As DataTable
    Dim vTable As DataTable = Nothing
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      vList("WebPageNumber") = pPageNumber
      vTable = DataHelper.GetWebPageInfo(vList)
    Catch vException As Exception
      ProcessError(vException)
    End Try
    Return vTable
  End Function

  Private Sub GetPageContent(ByVal pPageNumber As Integer, ByVal pLoginPageNumber As Integer, ByVal pUpdateDetailsPageNumber As Integer)
    Try
      mvCareControls = New List(Of CareWebControl)
      Dim vList As New ParameterList(HttpContext.Current)
      vList("WebPageNumber") = pPageNumber
      Dim vTable As DataTable = DataHelper.GetWebPageItems(vList)
      Dim vCurrentRow As HtmlTableRow = RowData
      HeadingData.Visible = False
      LeftData.Visible = False
      CenterData.Visible = False
      RightData.Visible = False
      FootingData.Visible = False
      Dim vLeftCell As HtmlTableCell = LeftData
      Dim vRightCell As HtmlTableCell = RightData
      Dim vCenterCell As HtmlTableCell = CenterData
      Dim vNextDataRow As Integer = 2
      If vTable IsNot Nothing Then
        For Each vRow As DataRow In vTable.Rows
          'Dim vPlaceHolder As PlaceHolder = Nothing
          Dim vHTMLCell As HtmlTableCell = Nothing
          Dim vContentPosition As DataDisplay.ContentPosition
          Select Case vRow("WebPageItemType").ToString
            Case "L"
              'vPlaceHolder = phcLeft
              vHTMLCell = vLeftCell
              vContentPosition = DataDisplay.ContentPosition.cpLeft
              vLeftCell.Visible = True
            Case "R"
              'vPlaceHolder = phcRight
              vHTMLCell = vRightCell
              vContentPosition = DataDisplay.ContentPosition.cpRight
              vRightCell.Visible = True
            Case "C"
              'vPlaceHolder = phcCenter
              vHTMLCell = vCenterCell
              vContentPosition = DataDisplay.ContentPosition.cpCenter
              vCenterCell.Visible = True
            Case "H"
              'vPlaceHolder = phcHeading
              vHTMLCell = HeadingData
              vContentPosition = DataDisplay.ContentPosition.cpHeader
              HeadingData.Visible = True
            Case "F"
              'vPlaceHolder = phcFooting
              vHTMLCell = FootingData
              vContentPosition = DataDisplay.ContentPosition.cpFooter
              FootingData.Visible = True
            Case "W"
              vCurrentRow = New HtmlTableRow
              vHTMLCell = New HtmlTableCell
              vHTMLCell.ColSpan = 3
              vHTMLCell.Visible = True
              vCurrentRow.Cells.Add(vHTMLCell)
              CenterTable.Rows.Insert(vNextDataRow, vCurrentRow)
              vNextDataRow += 1
              vCurrentRow = New HtmlTableRow
              vLeftCell = New HtmlTableCell
              vLeftCell.Attributes.CssStyle.Value = "LeftCell"
              vLeftCell.Visible = False
              vCurrentRow.Cells.Add(vLeftCell)
              vCenterCell = New HtmlTableCell
              vCenterCell.Attributes.CssStyle.Value = "CenterCell"
              vCenterCell.Visible = False
              vCurrentRow.Cells.Add(vCenterCell)
              vRightCell = New HtmlTableCell
              vRightCell.Attributes.CssStyle.Value = "RightCell"
              vRightCell.Visible = False
              vCurrentRow.Cells.Add(vRightCell)
              CenterTable.Rows.Insert(vNextDataRow, vCurrentRow)
              vNextDataRow += 1
          End Select
          If vHTMLCell IsNot Nothing Then
            Dim vControlPath As String = vRow.Item("ControlPath").ToString
            If vControlPath.Length > 0 Then
              Dim vControl As System.Web.UI.Control = LoadControl(vControlPath)
              vControl.ID = "ItemNumber" & vRow.Item("WebPageItemNumber").ToString
              Dim vCC As CareWebControl = TryCast(vControl, CareWebControl)
              If vCC IsNot Nothing Then
                mvCareControls.Add(vCC)
                With vCC
                  .PageCareControls = mvCareControls
                  .WebPageNumber = pPageNumber
                  .WebPageItemNumber = IntegerValue(vRow("WebPageItemNumber").ToString)
                  .WebPageItemName = vRow("WebPageItemName").ToString
                  .GroupName = vRow("ItemGroupName").ToString
                  .ParentGroup = vRow("ParentGroupName").ToString
                  .Position = vContentPosition
                  .HTML = vRow("WebPageHtml").ToString
                  .ItemStyle = vRow("WebPageItemStyle").ToString
                  .NumberOfRows = IntegerValue(vRow("NumberOfRows").ToString)
                  .LoginPageNumber = pLoginPageNumber
                  .UpdateDetailsPageNumber = pUpdateDetailsPageNumber
                  Dim vInitialParams As New ParameterList
                  vInitialParams.FillFromValueList(vRow("InitialParameters").ToString)
                  .InitialParameters = vInitialParams
                  .RequiredInitialParameters = vRow("RequiredInitialParameters").ToString
                  Dim vDefaultParams As New ParameterList
                  vDefaultParams.FillFromValueList(vRow("DefaultParameters").ToString)
                  .DefaultParameters = vDefaultParams
                  .RequiredDefaultParameters = vRow("RequiredDefaultParameters").ToString
                  .SubmitItemNumber = IntegerValue(vRow("SubmitItemNumber").ToString)
                  .SubmitItemUrl = vRow("SubmitURL").ToString
                  If vTable.Columns.Contains("WebPageUserControl") Then .SetAuthenticationRequired(vRow("WebPageUserControl").ToString)
                  If Request.QueryString("cwpd") <> "Y" Then
                    If .UseNewContact Then
                      If Session("ContactNumber") Is Nothing Then
                        Throw New CareException("Web Page is configured to use the contact details of a newly created contact but none was found")
                      End If
                    End If
                    If .NeedsAuthentication Then CheckAuthentication(False, pLoginPageNumber)
                  End If
                End With
              End If
              vHTMLCell.Controls.Add(vControl)
              If vCC IsNot Nothing Then
                If vCC.CenterControl Then vHTMLCell.Align = "Center"
              End If
            End If
          End If
        Next
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub ProcessError(ByVal pException As Exception)
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vResult As DataTable
    Dim vErrorId As Integer = 0
    Dim vErrorURL As String = "ShowErrors.aspx"
    LogException(pException)
    Session("LastException") = pException
    Try
      'Check Exception Type and fetch Error Number and Source if it is a CareException.
      If TypeOf (pException) Is CareException Then
        Dim vEx As CareException = CType(pException, CareException)
        vList("ErrorNumber") = vEx.ErrorNumber
        vList("ErrorSource") = vEx.Source
      Else
        vList("ErrorNumber") = 0
        vList("ErrorSource") = pException.Source
      End If
      vList("WebPageNumber") = Request.QueryString("PN")
      vList("ErrorMessage") = pException.Message
      vList("StackTrace") = pException.StackTrace
      'Record Error in Database.
      vResult = DataHelper.AddErrorLog(vList)
      If vResult.Columns.Contains("ErrorId") AndAlso vResult.Rows(0).Item("ErrorId").ToString.Length > 0 Then
        vErrorId = IntegerValue(vResult.Rows(0).Item("ErrorId").ToString)
        vErrorURL = String.Format(vErrorURL & "?EI={0}", vErrorId)
      End If
      ProcessRedirect(vErrorURL)
    Catch vEx As ThreadAbortException
      'Do nothing since it is expected from ProcessRedirect
    Catch ex As Exception
      ProcessRedirect("ShowErrors.aspx")
    End Try
  End Sub

  Private Function InWebPageDesigner() As Boolean
    If Request.QueryString("cwpd") = "Y" Then Return True
  End Function

  Private Sub ProcessRedirect(ByVal pPage As String)
    ' we cannot call Server.Transfer on an ASP.NET AJAX enabled page
    ' It throws error, thats why used Response.Redirect
    ' Server.Transfer(pPage, False)
    RedirectViaWhiteList(pPage)
  End Sub

  Private Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LoadComplete
    If Me.ScriptManager1 IsNot Nothing AndAlso Request("LastControl") IsNot Nothing AndAlso Request("LastControl").Length > 0 Then
      'Do nothing. This code is executed before the (focus) JavaScript but in .NET 4.0, the system remembers the control we set
      'the focus on (below) even after executing the (focus) JavaScript and hence the focus always remain on this control which is wrong.
      'In ScripResource.axd, the call to this._endPostBack(null, response, data); ends up calling the (focus) JavaScript but later
      'this._controlIDToFocus is checked which, if set (e.g. by calling vControl.FocusControl.Focus), changes the focus.
    ElseIf mvCareControls IsNot Nothing Then
      For Each vControl As CareWebControl In mvCareControls
        If vControl.FocusControl IsNot Nothing Then
          vControl.FocusControl.Focus()
          Exit For
        End If
      Next
    End If
    If Request.QueryString("pn") IsNot Nothing Then Session("LastPageNumber") = Request.QueryString("pn")
    If InWebPageDesigner() Then Response.CacheControl = "no-cache" 'always get the latest page for WPD
  End Sub

End Class

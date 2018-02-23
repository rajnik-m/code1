Public Class DisplayRelatedContacts
  Inherits CareWebControl
  Private mvOrganisationNumber As String
  Private mvIsButtonVisible As Boolean
  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vIsControlVisible As Boolean
      If Request.QueryString("ERCP") Is Nothing AndAlso Request.QueryString("MRCP") Is Nothing Then
        InitialiseControls(CareNetServices.WebControlTypes.wctRelatedContacts, tblDataEntry)
        If InitialParameters.OptionalValue("AccessView").Length > 0 Then
          If Not FindViewsOfUser("AccessView") Then
            Me.FindControl("PageError").Visible = True
            Me.FindControl("WarningMessage1").Visible = False
            Me.FindControl("WarningMessage2").Visible = False
            DirectCast(Me.FindControl("SendEmail"), Button).Visible = False
            DirectCast(Me.FindControl("MailMerge"), Button).Visible = False
            DirectCast(Me.FindControl("DataExport"), Button).Visible = False
            DirectCast(Me.FindControl("SetDefault"), Button).Visible = False
            vIsControlVisible = True
          Else
            Me.FindControl("PageError").Visible = False
          End If
        Else
          Me.FindControl("PageError").Visible = False
        End If
        If Not vIsControlVisible Then
          ShowButtons()
          BindDataGrid()
        End If
      ElseIf Request.QueryString("ERCP") IsNot Nothing AndAlso Request.QueryString("ERCP").Length > 0 AndAlso InitialParameters.OptionalValue("EditRelatedContactPage").Length > 0 Then
        Session("SelectedContactNumber") = Request.QueryString("CN")
        ProcessRedirect("default.aspx?pn=" & InitialParameters("EditRelatedContactPage").ToString)
      ElseIf Request.QueryString("MRCP") IsNot Nothing AndAlso Request.QueryString("MRCP").Length > 0 AndAlso InitialParameters.OptionalValue("MoveRelatedContactPage").Length > 0 Then
        Session("SelectedContactPositionNumber") = Request.QueryString("CPN")
        ProcessRedirect("default.aspx?pn=" & InitialParameters("MoveRelatedContactPage").ToString)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub BindDataGrid()
    Dim vList As New ParameterList(HttpContext.Current)
    If Not InWebPageDesigner() Then
      If Session("SelectedOrganisationNumber") IsNot Nothing AndAlso Session("SelectedOrganisationNumber").ToString.Length > 0 Then
        mvOrganisationNumber = Session("SelectedOrganisationNumber").ToString
      Else
        Throw New Exception("Organisation number is missing")
      End If
      vList("OrganisationNumber") = mvOrganisationNumber
    End If
    If InitialParameters.ContainsKey("ContactsView") Then vList("ViewName") = InitialParameters("ContactsView").ToString
    vList("SystemColumns") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber
    Dim vDataGrid As DataGrid = TryCast(Me.FindControl("RelatedContactData"), DataGrid)
    If vDataGrid IsNot Nothing Then
      vDataGrid.Columns.Clear()
      Dim vResult As String = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebRelatedContacts, vList)
      Dim vResultRowCount As Integer = DataHelper.FillGrid(vResult, vDataGrid)
      If vResultRowCount > 0 Then
        Dim vEditColumn As New BoundColumn()
        Dim vMoveColumn As New BoundColumn()
        Dim vContactPos As Integer
        Dim vEmailAddressPos As Integer
        Dim vContactPositionNumberPos As Integer
        vEditColumn.HeaderText = ""
        vDataGrid.Columns.AddAt(1, vEditColumn)
        vMoveColumn.HeaderText = ""
        vDataGrid.Columns.AddAt(2, vMoveColumn)
        vDataGrid.DataBind()
        For vCount As Integer = 3 To vDataGrid.Columns.Count - 1
          Dim vBoundColumn As BoundColumn = DirectCast(vDataGrid.Columns(vCount), BoundColumn)
          If vBoundColumn.DataField = "ContactNumber" Then
            vContactPos = vCount
          ElseIf vBoundColumn.DataField = "EmailAddress" Then
            vEmailAddressPos = vCount
          ElseIf vBoundColumn.DataField = "ContactPositionNumber" Then
            vContactPositionNumberPos = vCount
          End If
        Next
        For vRow As Integer = 0 To vDataGrid.Items.Count - 1
          vDataGrid.Items(vRow).Cells(1).Text = "<a href='default.aspx?pn=" & WebPageNumber & "&CN=" & vDataGrid.Items(vRow).Cells(vContactPos).Text & "&ERCP=1'>Edit</a>"
          vDataGrid.Items(vRow).Cells(2).Text = "<a href='default.aspx?pn=" & WebPageNumber & "&CPN=" & vDataGrid.Items(vRow).Cells(vContactPositionNumberPos).Text & "&MRCP=1'>Move</a>"
          Dim sb As StringBuilder = New StringBuilder("<a href='mailto:")
          sb.Append(vDataGrid.Items(vRow).Cells(vEmailAddressPos).Text)
          sb.Append("'>")
          sb.Append(vDataGrid.Items(vRow).Cells(vEmailAddressPos).Text)
          sb.Append("</a>")
          vDataGrid.Items(vRow).Cells(vEmailAddressPos).Text = sb.ToString()
        Next
        If Not mvIsButtonVisible Then
          vDataGrid.Columns(0).Visible = False
        End If
        'If Edit page is set then Edit hyperlink will be visible
        If InitialParameters.OptionalValue("EditRelatedContactPage").Length <= 0 Then vDataGrid.Columns(1).Visible = False
        'If Move page is set then Move hyperlink will be visible
        If InitialParameters.OptionalValue("MoveRelatedContactPage").Length <= 0 Then vDataGrid.Columns(2).Visible = False
        If InitialParameters.OptionalValue("UpdateView").Length > 0 Then
          If Not FindViewsOfUser("UpdateView") Then
            vDataGrid.Columns(1).Visible = False
            vDataGrid.Columns(2).Visible = False
          End If
        End If
      End If
      'Warning message should be invisible
      Me.FindControl("WarningMessage1").Visible = False
      Me.FindControl("WarningMessage2").Visible = False
    End If
  End Sub
  Private Function FindViewsOfUser(ByVal pParam As String) As Boolean
    Dim vContactViews As String = String.Empty
    Dim vViews() As String
    Dim vFound As Boolean
    If HttpContext.Current.User.Identity.IsAuthenticated Then
      If Not TypeOf (HttpContext.Current.User.Identity) Is System.Security.Principal.WindowsIdentity Then
        Dim vIdentity As FormsIdentity = CType(HttpContext.Current.User.Identity, FormsIdentity)
        If vIdentity.Ticket.UserData.Length > 0 Then
          Dim vItems As String() = vIdentity.Ticket.UserData.Split("|"c)
          If vItems.Length > 4 Then 'Check if viewname exists in userdata
            'Split again to get a list of views that the user belongs to
            vContactViews = vItems(4).ToString()
          End If
        End If
      End If
    End If
    vViews = vContactViews.Split(CChar(","))
    For Each vView As String In vViews
      If vView.ToUpper = InitialParameters.OptionalValue(pParam) Then
        vFound = True
        Exit For
      Else
        vFound = False
      End If
    Next
    Return vFound
  End Function
  Private Sub ShowButtons()
    'Set button visible, invisible on the basis of module parameteres.
    If FindControlByName(Me, "SendEmail") IsNot Nothing Then
      If InitialParameters.OptionalValue("SendEmailPage").Length > 0 Then
        DirectCast(Me.FindControl("SendEmail"), Button).Visible = True
        mvIsButtonVisible = True
      Else
        DirectCast(Me.FindControl("SendEmail"), Button).Visible = False
      End If
    End If
    If FindControlByName(Me, "MailMerge") IsNot Nothing Then
      If InitialParameters.OptionalValue("MailMergePage").Length > 0 Then
        DirectCast(Me.FindControl("MailMerge"), Button).Visible = True
        mvIsButtonVisible = True
      Else
        DirectCast(Me.FindControl("MailMerge"), Button).Visible = False
      End If
    End If
    If FindControlByName(Me, "DataExport") IsNot Nothing Then
      If InitialParameters.OptionalValue("DataExportPage").Length > 0 Then
        DirectCast(Me.FindControl("DataExport"), Button).Visible = True
        mvIsButtonVisible = True
      Else
        DirectCast(Me.FindControl("DataExport"), Button).Visible = False
      End If
    End If
    If FindControlByName(Me, "SetDefault") IsNot Nothing Then
      If InitialParameters.OptionalValue("SetDefaultPage").Length > 0 Then
        DirectCast(Me.FindControl("SetDefault"), Button).Visible = True
        mvIsButtonVisible = True
      Else
        DirectCast(Me.FindControl("SetDefault"), Button).Visible = False
      End If
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If Not InWebPageDesigner() Then
        Dim vCheck As Boolean
        Dim vCount As Integer = 0
        Dim vEmail As String = ""
        Dim vEmailAddressPos As Integer
        Dim vContactPos As Integer
        Dim vEmailList As New StringBuilder
        Dim vContactList As New StringBuilder
        Dim vDataGrid As DataGrid = TryCast(Me.FindControl("RelatedContactData"), DataGrid)
        Dim vShowWarningMessage As Boolean
        Dim vEmails As New Dictionary(Of String, String)

        For vRow As Integer = 3 To vDataGrid.Columns.Count - 1
          Dim vBoundColumn As BoundColumn = DirectCast(vDataGrid.Columns(vRow), BoundColumn)
          If vBoundColumn.DataField = "EmailAddress" Then
            vEmailAddressPos = vRow
          ElseIf vBoundColumn.DataField = "ContactNumber" Then
            vContactPos = vRow
          End If
        Next
        For vRow As Integer = 0 To vDataGrid.Items.Count - 1
          If CType(vDataGrid.Items(vRow).Cells(0).Controls(0), CheckBox).Checked Then
            vCheck = True
            vCount = vCount + 1
            vEmail = vDataGrid.Items(vRow).Cells(vEmailAddressPos).Text
            vEmail = Mid(vEmail, vEmail.IndexOf(">") + 2)
            vEmail = vEmail.Replace("</a>", "")
            If Not vEmails.ContainsValue(vEmail) And vEmail <> "&nbsp;" Then
              vEmailList.Append(vEmail & ";")
              vEmails.Add("EmailAddress" & vCount, vEmail)
            End If
            vContactList.Append(vDataGrid.Items(vRow).Cells(vContactPos).Text & ",")
          End If
        Next
        If vCheck Then
          'If checkbox checked then no warning message will display.
          Me.FindControl("WarningMessage2").Visible = False
          Me.FindControl("WarningMessage1").Visible = False
        Else
          'At least one row must be selected
          Me.FindControl("WarningMessage2").Visible = True
          Me.FindControl("WarningMessage1").Visible = False
          vShowWarningMessage = True
        End If
        If vCount > 1 AndAlso CType(sender, Button).ID = "SetDefault" Then
          'One row and only one row must be selected
          Me.FindControl("WarningMessage1").Visible = True
          Me.FindControl("WarningMessage2").Visible = False
          vShowWarningMessage = True
        End If
        If Not vShowWarningMessage Then
          'Send Email 
          If CType(sender, Button).ID = "SendEmail" Then
            If vEmailList.Length > 1 Then vEmailList.Remove(vEmailList.Length - 1, 1)
            Session("SelectedEmailAddresses") = vEmailList
            ProcessRedirect("default.aspx?pn=" & InitialParameters("SendEmailPage").ToString)
          End If
          'Mail Merge
          If CType(sender, Button).ID = "MailMerge" Then
            If vContactList.Length > 1 Then vContactList.Remove(vContactList.Length - 1, 1)
            Session("SelectedContacts") = vContactList
            ProcessRedirect("default.aspx?pn=" & InitialParameters("MailMergePage").ToString)
          End If
          'Data Export
          If CType(sender, Button).ID = "DataExport" Then
            If vContactList.Length > 1 Then vContactList.Remove(vContactList.Length - 1, 1)
            Session("SelectedContacts") = vContactList
            ProcessRedirect("default.aspx?pn=" & InitialParameters("DataExportPage").ToString)
          End If
          'Set Default
          If CType(sender, Button).ID = "SetDefault" Then
            If vContactList.Length > 1 Then vContactList.Remove(vContactList.Length - 1, 1)
            Dim vList As New ParameterList(HttpContext.Current)
            vList("ContactNumber") = mvOrganisationNumber
            vList("DefaultContactNumber") = vContactList.ToString
            DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vList)
            BindDataGrid()
          End If
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
End Class
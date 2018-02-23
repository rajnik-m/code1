Public Class UpdatePosition
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Private mvContactPositionNumberIndex As Integer
  Private mvAddressNumberIndex As Integer
  Private mvOrganisationNumberIndex As Integer

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Private Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init

  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctUpdatePosition, tblDataEntry)
      ShowAddressFields(False)
      SetLabelText("PageError", String.Empty)
      If InitialParameters.ContainsKey("ShowGridWithFields") AndAlso BooleanValue(InitialParameters("ShowGridWithFields").ToString) Then
        Dim vControl As Control = FindControlByName(Me, "OrganisationAddress")
        If vControl IsNot Nothing Then
          SetParentParentVisible("OrganisationAddress", False)
          Me.FindControl("PageError").Visible = True
          SetLabelText("PageError", "The Organisation Address combo box cannot be used when the ‘Show Grid with fields’ option is set.")
        End If
      End If
      BindPositionsDataGrid()
      SetControlEnabled("Name", False)
      SetControlEnabled("Address", False)
      'New button functionality is not available if the New Page parameter is not specified
      If Not InitialParameters.ContainsKey("NewPositionPage") Then
        Dim vControl As Control = FindControlByName(Me, "New")
        If vControl IsNot Nothing Then vControl.Visible = False
      End If
      If Request.QueryString("ON") IsNot Nothing AndAlso Request.QueryString("AN") IsNot Nothing Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = Request.QueryString("ON")
        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, vList))
        If vRow IsNot Nothing Then SetTextBoxText("Name", vRow("ContactName").ToString)
        vList("AddressNumber") = Request.QueryString("AN")
        vRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vList))
        If vRow IsNot Nothing Then
          SetTextBoxText("Address", vRow("AddressLine").ToString)
          PopulateAddressCombo(IntegerValue(vList("ContactNumber").ToString), IntegerValue(vList("AddressNumber").ToString))
        End If

        If Session("MovePosition") IsNot Nothing AndAlso Session("MovePosition").ToString = "Y" Then
          SetTextBoxText("Started", TodaysDate)
        End If
        If mvSupportsMultiView Then
          If IsPostBack Then
            HandleMultiViewDisplay()
          Else
            mvMultiView.SetActiveView(mvView2)  'We have selected an organisation, always display the edit fields
          End If
        End If
      Else
        HandleMultiViewDisplay()
        If ViewState("PositionOrganisation") IsNot Nothing AndAlso ViewState("PositionAddress") IsNot Nothing Then
          PopulateAddressCombo(IntegerValue(ViewState("PositionOrganisation").ToString), IntegerValue(ViewState("PositionAddress").ToString))
        End If
        Session.Remove("MovePosition")
        Session.Remove("OldValidFrom")
        Session.Remove("OldOrganisationNumber")
        Session.Remove("OldAddressNumber")
        Session.Remove("ContactPositionNumber")
      End If
      If FindControlByName(Me, "Started") Is Nothing Then
        AddHiddenField("HiddenValidFrom")
      End If

    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub


  Private Sub BindPositionsDataGrid()
    Dim vDGR As DataGrid = CType(Me.FindControl("OrganisationPositionData"), DataGrid)
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vContactNumber As Long

    vList("SystemColumns") = "Y"
    vList.Add("WebPageItemNumber", Me.WebPageItemNumber)
    If InitialParameters.ContainsKey("HideHistoricalPositions") AndAlso InitialParameters("HideHistoricalPositions").ToString = "Y" Then vList("Current") = "Y"
    If InitialParameters.ContainsKey("OrganisationGroup") AndAlso InitialParameters("OrganisationGroup").ToString.Length > 0 Then vList("OrganisationGroup") = InitialParameters("OrganisationGroup").ToString
    vContactNumber = GetContactNumberFromParentGroup()
    vDGR.Columns.Clear()
    DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, vDGR, String.Empty, vContactNumber, , vList, 0, True)
    If InitialParameters.ContainsKey("NewPositionPage") AndAlso InitialParameters.ContainsKey("HyperlinkText") AndAlso vDGR.Columns(1).HeaderText <> "" Then
      Dim vColumn As New TemplateColumn
      vColumn.HeaderText = ""
      vColumn.ItemTemplate = New EditTemplate("EditColumn2", 0, True, InitialParameters("HyperlinkText").ToString, "", "Current LIKE 'Y*'", If(InitialParameters.OptionalValue("HyperlinkFormat") = "B", True, False))
      vDGR.Columns.AddAt(1, vColumn)
      vDGR.DataBind()
    End If
    mvContactPositionNumberIndex = GetDataGridItemIndex(vDGR, "ContactPositionNumber")
    mvAddressNumberIndex = GetDataGridItemIndex(vDGR, "AddressNumber")
    mvOrganisationNumberIndex = GetDataGridItemIndex(vDGR, "ContactNumber")
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Try
      Dim vCommandName As String = ""
      If TypeOf e.CommandSource Is LinkButton Then
        vCommandName = DirectCast(e.CommandSource, LinkButton).CommandName
      ElseIf TypeOf e.CommandSource Is Button Then
        vCommandName = DirectCast(e.CommandSource, Button).CommandName
      End If
      Select Case vCommandName
        Case "Edit"
          Dim vContactPositionNumber As String = DirectCast(e.Item.Cells(mvContactPositionNumberIndex).Controls(0), ITextControl).Text
          ViewState("ContactPositionNumber") = vContactPositionNumber
          Dim vAddressNumber As String = DirectCast(e.Item.Cells(mvAddressNumberIndex).Controls(0), ITextControl).Text
          ViewState("AddressNumber") = vAddressNumber
          Dim vOrganisationNumber As String = DirectCast(e.Item.Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text
          ViewState("OrganisationNumber") = vOrganisationNumber

          Dim vList As New ParameterList(HttpContext.Current)
          'Get Contact Position Details
          vList("ContactNumber") = GetContactNumberFromParentGroup.ToString
          vList("ContactPositionNumber") = vContactPositionNumber
          Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, vList))
          If vRow IsNot Nothing Then
            SetTextBoxText("Name", vRow("ContactName").ToString)
            SetTextBoxText("Address", vRow("AddressLine").ToString)
            SetTextBoxText("Position", vRow("Position").ToString)
            SetTextBoxText("Location", vRow("Location").ToString)
            SetDropDownText("PositionFunction", vRow("PositionFunction").ToString)
            SetDropDownText("PositionSeniority", vRow("PositionSeniority").ToString)
            SetTextBoxText("Started", vRow("ValidFrom").ToString)
            SetHiddenText("HiddenValidFrom", vRow("ValidFrom").ToString)
            SetTextBoxText("Finished", vRow("ValidTo").ToString)
            SetCheckBoxChecked("Mail", BooleanValue(vRow("Mail").ToString))

            ViewState("PositionOrganisation") = vRow("ContactNumber").ToString
            ViewState("PositionAddress") = vRow("AddressNumber").ToString
            PopulateAddressCombo(IntegerValue(vRow("ContactNumber").ToString), IntegerValue(vRow("AddressNumber").ToString))
          End If
        Case InitialParameters("HyperlinkText").ToString
          'Redirect to the new position page to search for/create an organisation
          Dim vUrl As String = String.Format("default.aspx?pn={0}&ReturnURL={1}", InitialParameters("NewPositionPage"), Request.Url)
          Session("MovePosition") = "Y"
          Session("ContactPositionNumber") = DirectCast(e.Item.Cells(mvContactPositionNumberIndex).Controls(0), ITextControl).Text
          Dim vGrid As DataGrid = CType(Me.FindControl("OrganisationPositionData"), DataGrid)
          Session("OldValidFrom") = DirectCast(e.Item.Cells(GetDataGridItemIndex(vGrid, "ValidFrom")).Controls(0), ITextControl).Text
          Session("OldOrganisationNumber") = DirectCast(e.Item.Cells(GetDataGridItemIndex(vGrid, "ContactNumber")).Controls(0), ITextControl).Text
          Session("OldAddressNumber") = DirectCast(e.Item.Cells(GetDataGridItemIndex(vGrid, "AddressNumber")).Controls(0), ITextControl).Text
          ProcessRedirect(vUrl)
      End Select
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub PopulateAddressCombo(pContactNumber As Integer, pAddressNumber As Integer)
    Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "OrganisationAddress"), DropDownList)
    If vDDL IsNot Nothing Then
      Dim vList As New ParameterList(Me.Context)
      vList("ContactNumber") = pContactNumber
      Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vList)
      If vTable IsNot Nothing Then
        If FindControlByName(Me, "Town") IsNot Nothing AndAlso FindControlByName(Me, "Address2") IsNot Nothing Then
          vTable.Rows.InsertAt(vTable.NewRow(), 0)
          vTable.Rows(0).Item("AddressLine") = "<New Address>"
          vTable.Rows(0).Item("AddressNumber") = "0"
          vTable.Rows(0).Item("Historical") = ""
        End If
        vTable.DefaultView.RowFilter = "Historical  = '' OR AddressNumber = " & pAddressNumber
        vDDL.Items.Clear()
        vDDL.SelectedIndex = -1
        vDDL.SelectedValue = Nothing
        vDDL.ClearSelection()
        vDDL.DataSource = vTable
        vDDL.DataValueField = "AddressNumber"
        vDDL.DataTextField = "AddressLine"
        vDDL.DataBind()
        If pAddressNumber > 0 Then SetDropDownText("OrganisationAddress", pAddressNumber.ToString)
        vDDL.AutoPostBack = True
        AddHandler vDDL.SelectedIndexChanged, AddressOf SelectedAddressChanged
      End If
    End If
  End Sub

  Private Sub SelectedAddressChanged(pSender As Object, e As EventArgs)
    Dim vDDL As DropDownList = TryCast(pSender, DropDownList)
    If vDDL IsNot Nothing Then
      If vDDL.SelectedIndex = 0 Then
        ShowAddressFields(True)
      Else
        ShowAddressFields(False)
      End If
    End If
  End Sub

  Private Sub ShowAddressFields(pVisible As Boolean)
    SetParentParentVisible("PostcoderPostcode", pVisible)
    If FindControlByName(Me, "PostcoderAddress") Is Nothing Then
      SetParentParentVisible("Address2", pVisible)
      SetParentParentVisible("Town", pVisible)
      SetParentParentVisible("County", pVisible)
      SetParentParentVisible("Postcode", pVisible)
      SetParentParentVisible("Country", pVisible)
    Else
      If FindControlByName(Me, "PostcoderAddress") IsNot Nothing Then FindControlByName(Me, "PostcoderAddress").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Address") IsNot Nothing Then FindControlByName(Me, "Address").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Town") IsNot Nothing Then FindControlByName(Me, "Town").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "County") IsNot Nothing Then FindControlByName(Me, "County").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Postcode") IsNot Nothing Then FindControlByName(Me, "Postcode").Parent.Parent.Parent.Parent.Visible = pVisible
      If FindControlByName(Me, "Country") IsNot Nothing Then FindControlByName(Me, "Country").Parent.Parent.Parent.Parent.Visible = pVisible
    End If
  End Sub

  Public Overrides Sub ProcessButtonClickEvent(ByVal pValues As Object)
    Try
      ' Retrieve values from object
      Dim vButtonID As String = TryCast(pValues, String)
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vResultList As ParameterList

      Select Case vButtonID
        Case "Save"
          vList("ContactNumber") = GetContactNumberFromParentGroup.ToString
          vList("ValidFrom") = Today.ToString(CAREDateFormat)
          If FindControlByName(Me, "Started") IsNot Nothing Then
            vList("ValidFrom") = GetTextBoxText("Started")
          ElseIf Not String.IsNullOrEmpty(GetHiddenText("HiddenValidFrom")) Then
            vList("ValidFrom") = GetHiddenText("HiddenValidFrom")
          End If
          If FindControlByName(Me, "Finished") IsNot Nothing Then
            vList("ValidTo") = GetTextBoxText("Finished")
          End If
          If FindControlByName(Me, "Mail") IsNot Nothing Then
            'BR20862 - If the Mail control is hidden do not allow changes to its value. Control only exists if it is visible in WPD.
            vList("Mail") = BooleanString(GetCheckBoxChecked("Mail"))
          End If
          vList("Position") = GetTextBoxText("Position")
          vList("PositionLocation") = GetTextBoxText("Location")
          vList("PositionFunction") = GetDropDownValue("PositionFunction")
          vList("PositionSeniority") = GetDropDownValue("PositionSeniority")
          vList("UserID") = UserContactNumber.ToString
          If Not String.IsNullOrEmpty(Convert.ToString(ViewState("ContactPositionNumber"))) AndAlso _
            Convert.ToString(ViewState("ContactPositionNumber")) <> "0" Then
            AddUserParameters(vList)
            vList("ContactPositionNumber") = Convert.ToString(ViewState("ContactPositionNumber"))
            'If the organisation address combobox is visible on the page and the selected item does not match the existing address then we have to move the contact
            Dim vAddressUpdate As Boolean
            Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "OrganisationAddress"), DropDownList)
            If vDDL IsNot Nothing Then
              If Not String.IsNullOrEmpty(ViewState("AddressNumber").ToString) AndAlso IntegerValue(ViewState("AddressNumber").ToString) > 0 AndAlso _
                 Not String.IsNullOrEmpty(ViewState("OrganisationNumber").ToString) AndAlso IntegerValue(ViewState("OrganisationNumber").ToString) > 0 Then
                Dim vNewAddressNumber As Integer = IntegerValue(vDDL.SelectedValue)
                Dim vOldAddressNumber As Integer = IntegerValue(ViewState("AddressNumber").ToString)
                Dim vOrganisationNumber As Integer = IntegerValue(ViewState("OrganisationNumber").ToString)
                If vOldAddressNumber <> vNewAddressNumber Then
                  vAddressUpdate = True
                  If vNewAddressNumber = 0 Then                    'This is a new address for the organisation
                    vNewAddressNumber = AddNewOganisationAddress(vOrganisationNumber)
                  End If
                  vList("OrganisationNumber") = vOrganisationNumber
                  vList("AddressNumber") = vNewAddressNumber
                  vList("OldValidTo") = Date.Today.AddDays(-1)
                  vList("ValidFrom") = Date.Today
                  vResultList = DataHelper.MovePosition(vList)
                  ViewState("ContactPositionNumber") = vResultList("ContactPositionNumber")
                End If
              End If
            End If
            If vAddressUpdate = False Then vResultList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctPosition, vList)
          ElseIf Session("MovePosition") IsNot Nothing AndAlso Session("MovePosition").ToString = "Y" Then
            Dim vOrganisationNumber As Integer = IntegerValue(Request.QueryString("ON"))
            Dim vNewAddressNumber As Integer = IntegerValue(Request.QueryString("AN"))
            Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "OrganisationAddress"), DropDownList)
            If vDDL IsNot Nothing Then
              vNewAddressNumber = IntegerValue(vDDL.SelectedValue)
              If vNewAddressNumber <> IntegerValue(Request.QueryString("AN")) Then
                If vNewAddressNumber = 0 Then                    'This is a new address for the organisation
                  vNewAddressNumber = AddNewOganisationAddress(vOrganisationNumber)
                End If
              End If
            End If
            AddUserParameters(vList)
            vList("ContactPositionNumber") = Session("ContactPositionNumber").ToString
            vList("OrganisationNumber") = vOrganisationNumber.ToString
            vList("AddressNumber") = vNewAddressNumber
            vList("OldValidTo") = Today.AddDays(-1).ToString(CAREDateFormat)
            Dim vMinLeave As Date
            If DateTime.TryParse(Session("OldValidFrom").ToString, vMinLeave) Then
              If Date.Parse(vList("OldValidTo").ToString) < vMinLeave Then
                vList("OldValidTo") = vMinLeave.ToString(AppValues.DateFormat)
              End If
            End If
            Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, IntegerValue(Session("OldOrganisationNumber").ToString))
            For Each vDataRow As DataRow In vTable.Rows
              If IntegerValue(vDataRow.Item("AddressNumber").ToString) = IntegerValue(Session("OldAddressNumber").ToString) Then
                Dim vMaxLeave As Date
                If Date.TryParse(vDataRow.Item("ValidTo").ToString, vMaxLeave) Then
                  If Date.Parse(vList("OldValidTo").ToString) > vMaxLeave Then
                    vList("OldValidTo") = vMaxLeave.ToString(AppValues.DateFormat)
                  End If
                End If
                Exit For
              End If
            Next
            vResultList = DataHelper.MovePosition(vList)
            ViewState("ContactPositionNumber") = vResultList("ContactPositionNumber")
            Session.Remove("MovePosition")
            Session.Remove("OldValidFrom")
            Session.Remove("OldOrganisationNumber")
            Session.Remove("OldAddressNumber")
            Session.Remove("ContactPositionNumber")
          Else
            'Check if the org number and address number are present in the querystring and create a new position
            If Request.QueryString("ON") IsNot Nothing AndAlso Request.QueryString("AN") IsNot Nothing Then
              Dim vOrganisationNumber As Integer = IntegerValue(Request.QueryString("ON"))
              Dim vNewAddressNumber As Integer = IntegerValue(Request.QueryString("AN"))
              Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "OrganisationAddress"), DropDownList)
              If vDDL IsNot Nothing Then
                vNewAddressNumber = IntegerValue(vDDL.SelectedValue)
                If vNewAddressNumber <> IntegerValue(Request.QueryString("AN")) Then
                  If vNewAddressNumber = 0 Then                    'This is a new address for the organisation
                    vNewAddressNumber = AddNewOganisationAddress(vOrganisationNumber)
                  End If
                End If
              End If
              vList("OrganisationNumber") = vOrganisationNumber
              vList("AddressNumber") = vNewAddressNumber
              vResultList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctPosition, vList)
              ViewState("ContactPositionNumber") = vResultList("ContactPositionNumber")
            End If
          End If
          BindPositionsDataGrid()
          HighLightDataGridRow(Convert.ToString(ViewState("ContactPositionNumber")))
        Case "New", "GridHyperlink"
          'Redirect to the new position page to search for/create an organisation
          Session.Remove("MovePosition")
          Session.Remove("OldValidFrom")
          Session.Remove("OldOrganisationNumber")
          Session.Remove("OldAddressNumber")
          Session.Remove("ContactPositionNumber")
          Dim vUrl As String = String.Format("default.aspx?pn={0}&ReturnURL={1}", InitialParameters("NewPositionPage"), Request.Url)
          ProcessRedirect(vUrl)
        Case "Cancel"
          ViewState.Remove("PositionOrganisation")
          ViewState.Remove("PositionAddress")
      End Select
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enPositionDatesExceedSiteDates Then
        SetLabelText("PageError", vEx.Message)
        If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
      Else
        ProcessError(vEx)
      End If
    End Try
  End Sub

  Private Function AddNewOganisationAddress(pOrganisationNumber As Integer) As Integer
    'This is a new address for the organisation
    Dim vNewAddressParams As New ParameterList(HttpContext.Current)
    vNewAddressParams("ContactNumber") = pOrganisationNumber
    Dim vCheckExists As Boolean = True
    Dim vAddressValue As String = GetTextBoxText("Address2", vCheckExists)
    If vCheckExists Then vNewAddressParams("Address") = vAddressValue
    AddOptionalTextBoxValue(vNewAddressParams, "Town")
    AddOptionalTextBoxValue(vNewAddressParams, "County", True)
    If FindControlByName(Me, "Postcode") IsNot Nothing Then
      AddOptionalTextBoxValue(vNewAddressParams, "Postcode", True)
    Else
      Dim vValue As String = GetTextBoxText("PostcoderPostcode")
      If vValue.Length > 0 Then vNewAddressParams("Postcode") = vValue
    End If
    Dim vCountry As String = GetDropDownValue("Country")
    If vCountry.Length = 0 Then vCountry = "UK" 'Only add default country when adding a new contact
    If vCountry.Length > 0 Then vNewAddressParams("Country") = vCountry
    Dim vResultList As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctAddresses, vNewAddressParams)
    Return IntegerValue(vResultList("AddressNumber").ToString)
  End Function


  Private Sub HighLightDataGridRow(ByVal pContactPositionNumber As String)
    Try
      Dim vDGR As DataGrid = TryCast(Me.FindControl("OrganisationPositionData"), DataGrid)
      If Not vDGR Is Nothing Then
        For vCount As Integer = 0 To vDGR.Items.Count - 1
          If vDGR.Items(vCount).Cells(mvContactPositionNumberIndex).Text = pContactPositionNumber Then
            vDGR.SelectedIndex = vDGR.Items(vCount).ItemIndex
          End If
        Next
      End If
    Catch vException As Exception
      Throw vException
    End Try
  End Sub

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    'New hyper link should be hidden if the New button is not displayed
    Return If(Not InitialParameters.ContainsKey("NewPositionPage") OrElse FindControlByName(mvView2, "New") Is Nothing, False, True)
  End Function

End Class
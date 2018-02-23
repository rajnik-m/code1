Partial Public Class UpdateAddress
  Inherits CareWebControl
  Implements ICareChildWebControl
  Implements IMultiViewWebControl

  Private mvAddressNumberIndex As Integer
  Private mvDefaultIndex As Integer

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctUpdateAddress, tblDisplayData)
      BindAddressDataGrid()
      SetControlEnabled("BuildingNumber", False)
      ' Enable Default Button
      SetControlEnabled("Default", True)
      HandleMultiViewDisplay("Default")
      'SetLabelText("Message", String.Empty)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub BindAddressDataGrid()
    Try
      Dim vDGR As DataGrid = CType(Me.FindControl("ContactAddress"), DataGrid)
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vRestriction As String = ""
      Dim vContactNumber As Long

      vList("SystemColumns") = "Y"
      vContactNumber = GetContactNumberFromParentGroup()
      vList.Add("WebPageItemNumber", Me.WebPageItemNumber)
      If InitialParameters.ContainsKey("HideHistoricalAddresses") AndAlso InitialParameters("HideHistoricalAddresses").ToString.Length > 0 Then
        If InitialParameters("HideHistoricalAddresses").ToString = "Y" Then
          vRestriction = "Historical <> 'Yes'"
        End If
      End If
      vDGR.Columns.Clear()
      DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vDGR, vRestriction, vContactNumber, , vList, 0, If(InitialParameters.ContainsKey("HyperlinkText2") AndAlso InitialParameters.ContainsKey("HyperlinkText1") = False, False, True), If(InitialParameters.ContainsKey("HyperlinkText1"), InitialParameters("HyperlinkText1").ToString, "Edit"))
      If InitialParameters.ContainsKey("HyperlinkText2") AndAlso InitialParameters("HyperlinkText2").ToString.Length > 0 AndAlso vDGR.Columns(1).HeaderText <> "" Then
        Dim vColumn As New TemplateColumn
        vColumn.HeaderText = ""
        vColumn.ItemTemplate = New EditTemplate("EditColumn2", 0, True, InitialParameters("HyperlinkText2").ToString, "", "Default NOT LIKE 'Y*' AND Historical NOT LIKE 'Y*'", False)
        vDGR.Columns.AddAt(If(TypeOf vDGR.Columns(0) Is BoundColumn, 0, 1), vColumn)
        vDGR.DataBind()
      End If
    Catch vException As Exception
      Throw vException
    End Try
  End Sub

  Private Sub HighLightDataGridRow(ByVal vAddressNumber As String)
    Try
      Dim vDGR As DataGrid = TryCast(Me.FindControl("ContactAddress"), DataGrid)
      If Not vDGR Is Nothing Then
        For vCount As Integer = 0 To vDGR.Items.Count - 1
          If vDGR.Items(vCount).Cells(GetDataGridItemIndex(vDGR, "AddressNumber")).Text = vAddressNumber Then
            vDGR.SelectedIndex = vDGR.Items(vCount).ItemIndex
          End If
        Next
      End If
    Catch vException As Exception
      Throw vException
    End Try
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    'Nothing to do as this is a display only control
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = GetContactNumberFromParentGroup().ToString
      Dim vDGR As DataGrid = CType(Me.FindControl("ContactAddress"), DataGrid)
      ViewState("AddressNumber") = DirectCast(e.Item.Cells(GetDataGridItemIndex(vDGR, "AddressNumber")).Controls(0), ITextControl).Text
      vList("AddressNumber") = ViewState("AddressNumber")

      If e.CommandName = If(InitialParameters.ContainsKey("HyperlinkText1"), InitialParameters("HyperlinkText1").ToString, "Edit") Then
        ViewState("Default") = DirectCast(e.Item.Cells(GetDataGridItemIndex(vDGR, "Default")).Controls(0), ITextControl).Text

        ' Get Main Contact Information
        Dim vContactType As ContactInfo.ContactTypes
        Dim vEnable As Boolean = True
        Dim vRowContactType As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, vList))
        If vRowContactType IsNot Nothing Then
          Select Case vRowContactType("ContactType").ToString
            Case "O"
              vContactType = ContactInfo.ContactTypes.ctOrganisation
            Case "J"
              vContactType = ContactInfo.ContactTypes.ctJoint
            Case Else
              vContactType = ContactInfo.ContactTypes.ctContact
          End Select
        End If
        ' Get Address Details
        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vList))
        If vRow IsNot Nothing Then
          SetDropDownText("Country", vRow("CountryCode").ToString)
          SetTextBoxText("Postcode", vRow("Postcode").ToString)
          SetTextBoxText("BuildingNumber", vRow("BuildingNumber").ToString)
          SetTextBoxText("HouseName", vRow("HouseName").ToString)
          SetTextBoxText("Address", vRow("Address").ToString)
          SetTextBoxText("Town", vRow("Town").ToString)
          SetTextBoxText("County", vRow("County").ToString)
          SetDropDownText("Branch", vRow("Branch").ToString)
          SetTextBoxText("ValidFrom", vRow("ValidFrom").ToString)
          SetTextBoxText("ValidTo", vRow("ValidTo").ToString)
          SetLabelText("PafStatus", vRow.Item("Paf").ToString)
          ViewState("PafStatus") = vRow.Item("Paf").ToString
          ' Enable/Disable Default Button
          SetControlEnabled("Default", Not vRow("Historical").ToString.StartsWith("Y"))
          If vRow.Item("AddressType").ToString = "O" And vContactType <> ContactInfo.ContactTypes.ctOrganisation Then vEnable = False
        End If
        EnableControls(vEnable)
      ElseIf e.CommandName = If(InitialParameters.ContainsKey("HyperlinkText2"), InitialParameters("HyperlinkText2").ToString, "") Then
        vList("Default") = "Y"
        DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctAddresses, vList)
        If TypeOf vDGR.Columns(1) Is BoundColumn Then GoToSubmitPage()
        BindAddressDataGrid()
        HighLightDataGridRow(Convert.ToString(ViewState("AddressNumber")))
        If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView1)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
      'SetErrorLabel(vException.Message)
    End Try
  End Sub

  Public Overrides Sub ProcessButtonClickEvent(ByVal pValues As Object)
    Try
      ' Retrieve values from object
      Dim vButtonID As String = TryCast(pValues, String)
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vResultList As ParameterList

      Select Case vButtonID
        Case "Save", "Default"
          Page.Validate()
          If IsValid() Then
            vList("ContactNumber") = GetContactNumberFromParentGroup().ToString
            AddUserParameters(vList)
            If Not String.IsNullOrEmpty(Convert.ToString(ViewState("AddressNumber"))) AndAlso _
              Convert.ToString(ViewState("AddressNumber")) <> "0" Then
              vList("AddressNumber") = Convert.ToString(ViewState("AddressNumber"))
              ' Load Default Values
              Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, vList))
              If vRow IsNot Nothing Then
                vList("Country") = vRow("CountryCode").ToString
                vList("Postcode") = vRow("Postcode").ToString
                vList("BuildingNumber") = vRow("BuildingNumber").ToString
                vList("HouseName") = vRow("HouseName").ToString
                vList("Address") = vRow("Address").ToString
                vList("Town") = vRow("Town").ToString
                vList("County") = vRow("County").ToString
                vList("Branch") = vRow("Branch").ToString
                vList("ValidFrom") = vRow("ValidFrom").ToString
                vList("ValidTo") = vRow("ValidTo").ToString
                If Not String.IsNullOrEmpty(Convert.ToString(ViewState("Default"))) Then
                  If Convert.ToString(ViewState("Default")).ToUpper() = "YES" Then
                    vList("Default") = "Y"
                  End If
                End If
              End If
            Else
              ViewState("AddressNumber") = "0"
              ViewState("PafStatus") = ""
            End If
            vList("Country") = GetDropDownValue("Country")
            vList("Postcode") = GetTextBoxText("Postcode")
            vList("BuildingNumber") = GetTextBoxText("BuildingNumber")
            vList("HouseName") = GetTextBoxText("HouseName")
            vList("Address") = GetTextBoxText("Address")
            vList("Town") = GetTextBoxText("Town")
            If Not vList.Contains("Town") OrElse vList("Town").ToString.Length = 0 Then vList("Town") = "#"
            If vList("Town").ToString = "#" AndAlso vList("Country").ToString = "UK" Then vList.Remove("Town")
            vList("County") = GetTextBoxText("County")
            vList("Branch") = GetDropDownValue("Branch")
            vList("ValidFrom") = GetTextBoxText("ValidFrom")
            vList("ValidTo") = GetTextBoxText("ValidTo")
            ' If New Address is created, set address as default
            If vButtonID = "Default" Then vList("Default") = "Y"
            If ViewState("PafStatus") IsNot Nothing Then vList("PafStatus") = ViewState("PafStatus").ToString
            If Not vList.Contains("AddressNumber") Then
              vResultList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctAddresses, vList)
              ViewState("AddressNumber") = vResultList("AddressNumber")
            Else
              vResultList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctAddresses, vList)
            End If
            BindAddressDataGrid()
            HighLightDataGridRow(Convert.ToString(ViewState("AddressNumber")))
            'SetLabelText("Message", "Address saved as default")
            'Else
            '  SetLabelText("Message", "Address saved successfully")
          End If
        Case "New", "GridHyperlink"
          EnableControls(True)
          MyBase.ClearControls()
          Dim vDGR As DataGrid = CType(Me.FindControl("ContactAddress"), DataGrid)
          vDGR.SelectedIndex = -1
          ViewState("AddressNumber") = "0"
          ViewState("PafStatus") = ""

          ' Set Default Values
          SetDropDownText("Country", AppValues.DefaultCountryCode)
          SetTextBoxText("ValidFrom", AppValues.TodaysDate)
          SetTextBoxText("ValidTo", "")
      End Select

    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
      'SetErrorLabel(vException.Message)
    End Try
  End Sub

  Private Sub EnableControls(ByVal pEnable As Boolean)
    SetControlEnabled("PostcoderPostcode", pEnable)
    SetControlEnabled("PostcoderAddress", pEnable)

    SetControlEnabled("Country", pEnable)
    SetControlEnabled("Postcode", pEnable)
    SetControlEnabled("HouseName", pEnable)
    SetControlEnabled("Address", pEnable)
    SetControlEnabled("Town", pEnable)
    SetControlEnabled("County", pEnable)
    SetControlEnabled("Branch", pEnable)
    SetControlEnabled("ValidFrom", pEnable)
    SetControlEnabled("cmdFind" & "ValidFrom", pEnable)
    SetControlEnabled("ValidTo", pEnable)
    SetControlEnabled("cmdFind" & "ValidTo", pEnable)
  End Sub

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    'New hyper link should be hidden if the New button is not displayed
    Return If(FindControlByName(mvView2, "New") Is Nothing, False, True)
  End Function

End Class
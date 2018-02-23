Partial Public Class UpdateEmailAddress
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctUpdateEmailAddress, tblDisplayData)
      BindDataGrid()
      SetDropDownText("Device", "")
      HandleMultiViewDisplay()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub BindDataGrid()
    Try
      Dim vDGR As DataGrid = CType(Me.FindControl("EmailAddress"), DataGrid)
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vRestriction As String = ""
      vList("ContactNumber") = GetContactNumberFromParentGroup.ToString
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber

      vRestriction = "IsOrganisation = '' AND (Email = 'Y' OR WwwAddress = 'Y')"
      Dim vDevices As String = GetDevices()
      If vDevices.Length > 0 Then vRestriction &= " AND DeviceCode IN ( " & vDevices & ")"
      If InitialParameters.ContainsKey("HideHistoricalEmailWebAddresses") AndAlso InitialParameters("HideHistoricalEmailWebAddresses").ToString = "Y" Then
        vRestriction &= " AND IsActive = 'Yes'"
      End If
      Dim vResult As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, vList)
      DataHelper.FillGrid(vResult, vDGR, vRestriction, , If(InitialParameters.ContainsKey("HyperlinkText2") AndAlso InitialParameters.ContainsKey("HyperlinkText1") = False, False, True), If(InitialParameters.ContainsKey("HyperlinkText1"), InitialParameters("HyperlinkText1").ToString, "Edit"))
      If InitialParameters.ContainsKey("HyperlinkText2") AndAlso InitialParameters("HyperlinkText2").ToString.Length > 0 AndAlso vDGR.Columns(1).HeaderText <> "" Then
        Dim vColumn As New TemplateColumn
        vColumn.HeaderText = ""
        vColumn.ItemTemplate = New EditTemplate("EditColumn2", 0, True, InitialParameters("HyperlinkText2").ToString, "", "DeviceDefault NOT LIKE 'Y*' AND IsActive LIKE 'Y*'", False)
        vDGR.Columns.AddAt(If(TypeOf vDGR.Columns(0) Is BoundColumn, 0, 1), vColumn)
        vDGR.DataBind()
      End If
    Catch vException As Exception
      Throw vException
    End Try
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Dim vDGR As DataGrid = CType(Me.FindControl("EmailAddress"), DataGrid)
    Dim vCommunicationNumber As String = DirectCast(e.Item.Cells(GetDataGridItemIndex(vDGR, "CommunicationNumber")).Controls(0), ITextControl).Text
    'ClearControls(True)
    Dim vList As New ParameterList(HttpContext.Current)
    vList("ContactNumber") = GetContactNumberFromParentGroup.ToString
    vList("CommunicationNumber") = vCommunicationNumber
    ViewState("CommunicationNumber") = vCommunicationNumber

    If e.CommandName = If(InitialParameters.ContainsKey("HyperlinkText1"), InitialParameters("HyperlinkText1").ToString, "Edit") Then
      Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, vList)
      If Not vTable Is Nothing Then
        Dim vRow As DataRow = vTable.Rows(0)
        SetDropDownText("AddressNumber", vRow("AddressNumber").ToString)
        SetDropDownText("Device", vRow("DeviceCode").ToString, True)
        SetDropDownText("CommunicationUsage", vRow("CommunicationUsage").ToString)
        SetTextBoxText("Number", vRow("Number").ToString)
        If (String.Compare(vRow("Mail").ToString, "Yes", True) = 0) Then
          SetCheckBoxChecked("Mail")
        Else
          SetCheckBoxChecked("Mail", False)
        End If
        SetTextBoxText("ValidFrom", vRow("ValidFrom").ToString)
        If (String.Compare(vRow("DeviceDefault").ToString, "Yes", True) = 0) Then
          SetCheckBoxChecked("DeviceDefault")
        Else
          SetCheckBoxChecked("DeviceDefault", False)
        End If
        SetTextBoxText("ValidTo", vRow("ValidTo").ToString)
        If (String.Compare(vRow("PreferredMethod").ToString, "Yes", True) = 0) Then
          SetCheckBoxChecked("PreferredMethod")
        Else
          SetCheckBoxChecked("PreferredMethod", False)
        End If
        SetTextBoxText("Notes", vRow("Notes").ToString)
      End If
    ElseIf e.CommandName = If(InitialParameters.ContainsKey("HyperlinkText2"), InitialParameters("HyperlinkText2").ToString, "") Then
      AddUserParameters(vList)
      vList("OldContactNumber") = GetContactNumberFromParentGroup.ToString
      vList("DeviceDefault") = "Y"
      DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctNumber, vList)
      If TypeOf vDGR.Columns(1) Is BoundColumn Then GoToSubmitPage()
      BindDataGrid()
      HighLightDataGridRow(Convert.ToString(ViewState("CommunicationNumber")))
      If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView1)
    End If
  End Sub

  Public Overrides Sub ProcessButtonClickEvent(ByVal pValues As Object)
    Try
      Dim vButtonID As String = TryCast(pValues, String)
      Select Case vButtonID
        Case "Save", "Default"
          SaveEmailDetails()
        Case "New", "GridHyperlink"
          MyBase.ClearControls()
          Dim vDGR As DataGrid = CType(Me.FindControl("EmailAddress"), DataGrid)
          If Not vDGR Is Nothing Then
            vDGR.SelectedIndex = -1
          End If
          Dim vDDL As DropDownList = DirectCast(FindControlByName(Me, "AddressNumber"), DropDownList)
          If vDDL IsNot Nothing AndAlso vDDL.Items.Count > 0 Then
            vDDL.SelectedIndex = 1
          End If
          SetDropDownText("Device", "")
          SetTextBoxText("ValidFrom", CStr(Date.Now.Date))
          SetTextBoxText("ValidTo", "")
          ViewState("CommunicationNumber") = "0"
      End Select
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub SaveEmailDetails()
    Dim vResultList As ParameterList
    If (CType(ViewState("CommunicationNumber"), Integer) > 0) Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("ContactNumber") = GetContactNumberFromParentGroup.ToString
      vList("CommunicationNumber") = CType(ViewState("CommunicationNumber"), Integer)
      Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, vList)

      If Not vTable Is Nothing Then
        Dim vRow As DataRow = vTable.Rows(0)
        Dim vDevice As String = vRow("DeviceCode").ToString

        Dim vResult As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, vList)
        Dim vParams As New ParameterList(HttpContext.Current)
        vParams("ContactNumber") = GetContactNumberFromParentGroup.ToString
        AddUserParameters(vParams)
        vParams("OldContactNumber") = GetContactNumberFromParentGroup.ToString
        vParams("CommunicationNumber") = CType(ViewState("CommunicationNumber"), Integer)
        vParams("AddressNumber") = GetDropDownValue("AddressNumber")
        vParams("Device") = GetDropDownValue("Device")
        vParams("OldDevice") = vDevice
        vParams("DiallingCode") = GetDropDownValue("DiallingCode")
        vParams("STDCode") = GetDropDownValue("STDCode")
        vParams("Number") = GetTextBoxText("Number")
        vParams("ExDirectory") = BooleanString(GetCheckBoxChecked("ExDirectory"))
        vParams("Extension") = GetTextBoxText("Extension")
        vParams("Mail") = BooleanString(GetCheckBoxChecked("Mail"))
        vParams("PreferredMethod") = BooleanString(GetCheckBoxChecked("PreferredMethod"))
        vParams("DeviceDefault") = BooleanString(GetCheckBoxChecked("DeviceDefault"))
        vParams("ValidFrom") = GetTextBoxText("ValidFrom")
        vParams("ValidTo") = GetTextBoxText("ValidTo")
        vParams("Notes") = GetTextBoxText("Notes")
        Dim vExists As Boolean = True
        Dim vUsage As String = GetDropDownValue("CommunicationUsage", vExists)
        If vExists Then vParams("CommunicationUsage") = vUsage
        vResultList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctNumber, vParams)
      End If
    Else
      Dim vParams As New ParameterList(HttpContext.Current)
      vParams("ContactNumber") = GetContactNumberFromParentGroup.ToString
      If Convert.ToString(vParams("UserLogname")).Length = 0 Then
        vParams("UserID") = vParams("ContactNumber")
      End If
      vParams("AddressNumber") = GetDropDownValue("AddressNumber")
      vParams("Device") = GetDropDownValue("Device")
      vParams("DiallingCode") = GetDropDownValue("DiallingCode")
      vParams("STDCode") = GetDropDownValue("STDCode")
      vParams("Number") = GetTextBoxText("Number")
      vParams("ExDirectory") = BooleanString(GetCheckBoxChecked("ExDirectory"))
      vParams("Extension") = GetTextBoxText("Extension")
      vParams("Mail") = BooleanString(GetCheckBoxChecked("Mail"))
      vParams("PreferredMethod") = BooleanString(GetCheckBoxChecked("PreferredMethod"))
      vParams("ValidFrom") = GetTextBoxText("ValidFrom")
      vParams("ValidTo") = GetTextBoxText("ValidTo")
      vParams("Notes") = GetTextBoxText("Notes")
      Dim vExists As Boolean = True
      Dim vUsage As String = GetDropDownValue("CommunicationUsage", vExists)
      If vExists Then vParams("CommunicationUsage") = vUsage
      vResultList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctNumber, vParams)
      ViewState("CommunicationNumber") = vResultList("CommunicationNumber")
    End If
    BindDataGrid()
    HighLightDataGridRow(Convert.ToString(ViewState("CommunicationNumber")))
  End Sub

  Private Sub HighLightDataGridRow(ByVal vCommunicationNumber As String)
    Dim vDGR As DataGrid = CType(Me.FindControl("EmailAddress"), DataGrid)
    Dim vNumberIndex As Integer = GetDataGridItemIndex(vDGR, "CommunicationNumber")
    For count As Integer = 0 To vDGR.Items.Count - 1
      If vDGR.Items(count).Cells(vNumberIndex).Text = vCommunicationNumber Then
        vDGR.SelectedIndex = vDGR.Items(count).ItemIndex
      End If
    Next
  End Sub

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    'New hyper link should be hidden if the New button is not displayed
    Return If(FindControlByName(mvView2, "New") Is Nothing, False, True)
  End Function

End Class
Public Class UpdateCPDPoints
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Dim mvContactCPDCycleNumber As String
  Private mvCPDPointIndex As Integer
  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vCPDPeriod As DropDownList
      Dim vCpdCategoryType As DropDownList
      Dim vDataTable As DataTable
      Dim vResult As String

      InitialiseControls(CareNetServices.WebControlTypes.wctUpdateCpdPoints, tblDataEntry)
      If Not InWebPageDesigner() AndAlso Request.QueryString("CN") Is Nothing Then Throw New Exception("Parameter CycleNumber is missing")
      If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then
        mvContactCPDCycleNumber = Request.QueryString("CN")
      End If
      BindDataGrid()
      HandleMultiViewDisplay()
      If mvContactCPDCycleNumber IsNot Nothing Then
        vList("ContactNumber") = UserContactNumber()
        vList("ContactCpdCycleNumber") = mvContactCPDCycleNumber
        vList("SystemColumns") = "Y"

        vCPDPeriod = TryCast(Me.FindControl("ContactCpdPeriodNumber"), DropDownList)
        vCPDPeriod.DataTextField = "ContactCpdPeriodNumberDesc"
        vCPDPeriod.DataValueField = "ContactCpdPeriodNumber"
        DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDPeriods, vCPDPeriod, True, vList)
        vList.Remove("SystemColumns")

        vResult = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDCycles, vList)
        vDataTable = GetDataTable(vResult)
        If vDataTable.Rows.Count > 0 Then
          vCpdCategoryType = TryCast(Me.FindControl("CpdCategoryType"), DropDownList)
          vCpdCategoryType.DataTextField = "CpdCategoryTypeDesc"
          vCpdCategoryType.DataValueField = "CpdCategoryType"
          Dim vRestriction As String = "CPDCycleType ='" & vDataTable.Rows(0).Item("CPDCycleType").ToString & "' OR CPDCycleType Is Null"
          DataHelper.FillComboWithRestriction(CareNetServices.XMLLookupDataTypes.xldtCPDCategoryTypes, vCpdCategoryType, True, vList, vRestriction)
        End If
        SetTextBoxText("PointsDate", "")
        SetLabelText("PageError", "")
        SetLabelText("WarningMessage", "")
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try

  End Sub

  Private Sub BindDataGrid()
    Dim vList As New ParameterList(HttpContext.Current)
    If mvContactCPDCycleNumber IsNot Nothing Then vList("ContactCpdCycleNumber") = mvContactCPDCycleNumber
    vList("SystemColumns") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber
    vList("ContactNumber") = UserContactNumber()
    If InWebPageDesigner() Then
      vList("FromWPD") = "Y"
    End If
    vList("ForPortal") = "Y"
    Dim vDataGrid As DataGrid
    vDataGrid = TryCast(Me.FindControl("ContactCPDPoints"), DataGrid)
    If vDataGrid IsNot Nothing Then
      vDataGrid.Columns.Clear()
      Dim vDisplayEditColumn As Boolean = InitialParameters.ContainsKey("AllowEditing") AndAlso BooleanValue(InitialParameters("AllowEditing").ToString)
      Dim vHyperlinkText As String = ""
      Dim vEditColumnRestriction As String = ""
      If vDisplayEditColumn Then
        If InitialParameters.ContainsKey("HyperlinkText") AndAlso InitialParameters("HyperlinkText").ToString.Length > 0 Then
          vHyperlinkText = InitialParameters("HyperlinkText").ToString
        End If
        vEditColumnRestriction = "WebPublish = 'Y'"
      End If
      DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPoints, vDataGrid, "", UserContactNumber, , vList, 0, vDisplayEditColumn, vHyperlinkText, vEditColumnRestriction)
      mvCPDPointIndex = GetDataGridItemIndex(vDataGrid, "ContactCPDPointNumber")
    End If
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Try
      Dim vCPDPointNumber As String = DirectCast(e.Item.Cells(mvCPDPointIndex).Controls(0), ITextControl).Text
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vRow As DataRow
      Dim vDataTable As New DataTable
      ViewState("ContactCpdPointNumber") = vCPDPointNumber
      'Get Contact Cpd Point Details
      vList("ContactNumber") = UserContactNumber()
      vList("ContactCpdPointNumber") = vCPDPointNumber
      vList("ForPortal") = "Y"
      vRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDPoints, vList))
      If vRow IsNot Nothing Then
        TryCast(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).SelectedValue = vRow("ContactCpdPeriodNumber").ToString
        TryCast(Me.FindControl("CpdCategoryType"), DropDownList).SelectedValue = vRow("CategoryType").ToString
        TryCast(Me.FindControl("CpdCategory"), DropDownList).DataBind()
        SetDropDownText("CpdCategory", vRow("Category").ToString)
        vDataTable = CType(TryCast(Me.FindControl("CpdCategory"), DropDownList).DataSource, DataTable)
        Dim vRowColl() As DataRow = vDataTable.Select("CpdCategory = '" & vRow("Category").ToString & "'")
        If vRowColl(0).Item("PointsOverride").ToString = "Y" Then
          SetControlEnabled("CpdPoints", True)
          SetControlEnabled("CpdPoints2", True)
        Else
          SetControlEnabled("CpdPoints", False)
          SetControlEnabled("CpdPoints2", False)
        End If
        SetTextBoxText("PointsDate", vRow("PointsDate").ToString)
        SetTextBoxText("CpdPoints", vRow("Points").ToString)
        SetTextBoxText("Notes", vRow("Notes").ToString)
        If vRow("EvidenceSeen").ToString = "Y" Then
          SetCheckBoxChecked("EvidenceSeen", True)
        Else
          SetCheckBoxChecked("EvidenceSeen", False)
        End If
        SetLabelText("WarningMessage", "")
        SetTextBoxText("CpdPoints2", vRow("Points2").ToString)
        SetDropDownText("WebPublish", vRow("WebPublish").ToString)
        SetDropDownText("CpdItemType", vRow("ItemType").ToString)
        SetTextBoxText("CpdOutcome", vRow("Outcome").ToString)
      End If
      EnableControls(False)
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      If CType(sender, Control).ID = "Save" Then
        If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then
          If ViewState("DateMandatory") IsNot Nothing AndAlso ViewState("DateMandatory").ToString = "Y" AndAlso GetTextBoxText("PointsDate").Length = 0 Then
            SetLabelText("WarningMessage", "Points Date is mandatory can not be left blank")
            If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
          Else
            'Update CPD Point
            vList("UserID") = UserContactNumber()
            If ViewState("ContactCpdPointNumber") IsNot Nothing Then
              vList("ContactCpdPeriodNumber") = DirectCast(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).SelectedValue
              vList("ContactCpdPointNumber") = ViewState("ContactCpdPointNumber")
              vList("ContactNumber") = UserContactNumber()
              vList("CpdCategory") = DirectCast(Me.FindControl("CpdCategory"), DropDownList).SelectedValue
              vList("CpdCategoryType") = DirectCast(Me.FindControl("CpdCategoryType"), DropDownList).SelectedValue
              vList("CpdPoints") = GetTextBoxText("CpdPoints")
              vList("PointsDate") = GetTextBoxText("PointsDate")
              If DirectCast(Me.FindControl("EvidenceSeen"), CheckBox).Checked Then
                vList("EvidenceSeen") = "Y"
              Else
                vList("EvidenceSeen") = "N"
              End If
              vList("Notes") = GetTextBoxText("Notes")
              AddOptionalTextBoxValue(vList, "CpdPoints2", True)
              AddOptionalDropDownValue(vList, "WebPublish", True)
              AddOptionalDropDownValue(vList, "CpdItemType", True)
              AddOptionalTextBoxValue(vList, "CpdOutcome", True)
              DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctCPDPoints, vList)
              BindDataGrid()
              HighLightDataGridRow(ViewState("ContactCpdPointNumber").ToString)
            Else
              'Add CPD point
              vList("ContactCpdPeriodNumber") = DirectCast(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).SelectedValue
              vList("ContactNumber") = UserContactNumber()
              vList("CpdCategory") = DirectCast(Me.FindControl("CpdCategory"), DropDownList).SelectedValue
              vList("CpdCategoryType") = DirectCast(Me.FindControl("CpdCategoryType"), DropDownList).SelectedValue
              vList("CpdPoints") = GetTextBoxText("CpdPoints")
              vList("PointsDate") = GetTextBoxText("PointsDate")
              If DirectCast(Me.FindControl("EvidenceSeen"), CheckBox).Checked Then
                vList("EvidenceSeen") = "Y"
              Else
                vList("EvidenceSeen") = "N"
              End If
              vList("Notes") = GetTextBoxText("Notes")
              AddOptionalTextBoxValue(vList, "CpdPoints2", True)
              AddOptionalDropDownValue(vList, "WebPublish", True)
              AddOptionalDropDownValue(vList, "CpdItemType", True)
              AddOptionalTextBoxValue(vList, "CpdOutcome", True)
              Dim vReturnList As New ParameterList(HttpContext.Current)
              vReturnList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctCPDPoints, vList)
              BindDataGrid()
              HighLightDataGridRow(vReturnList.Item("ContactCpdPointNumber").ToString)
              ViewState("ContactCpdPointNumber") = vReturnList.Item("ContactCpdPointNumber").ToString
            End If
            SetLabelText("PageError", "")
            SetLabelText("WarningMessage", "")
          End If
        End If
      ElseIf (CType(sender, Control).ID = "New" OrElse CType(sender, Control).ID = "GridHyperlink") And Not InWebPageDesigner() Then
        'Clear the fields  SetDropDownText("CpdCategoryType", String.Empty, True)
        Dim vDropDown As DropDownList = DirectCast(Me.FindControl("CpdCategory"), DropDownList)
        vList("FromCPDPoints") = "Y"
        vDropDown.DataTextField = "CpdCategoryDesc"
        vDropDown.DataValueField = "CpdCategory"
        DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDCategories, vDropDown, True, vList)
        SetDropDownText("ContactCpdPeriodNumber", String.Empty)
        SetDropDownText("CpdCategoryType", String.Empty)
        SetDropDownText("CpdCategory", String.Empty)
        SetTextBoxText("PointsDate", "")
        SetTextBoxText("CpdPoints", String.Empty)
        SetTextBoxText("EvidenceSeen", String.Empty)
        SetTextBoxText("Notes", String.Empty)
        SetLabelText("PageError", String.Empty)
        SetLabelText("WarningMessage", String.Empty)
        SetControlEnabled("CpdPoints", True)
        SetControlEnabled("CpdPoints2", True)
        SetTextBoxText("CpdPoints2", "")
        SetDropDownText("WebPublish", "Y")
        SetDropDownText("CpdItemType", "")
        SetTextBoxText("CpdOutcome", "")
        EnableControls(True)
        ViewState.Clear()
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As CareException
      SetLabelText("WarningMessage", "")
      SetLabelText("PageError", vEx.Message)
      If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
    Catch vEx As Exception
      Throw vEx
    End Try
  End Sub

  Private Sub HighLightDataGridRow(ByVal pContactCPDCycleNumber As String)
    Try
      Dim vDGR As DataGrid = TryCast(Me.FindControl("ContactCPDPoints"), DataGrid)
      If Not vDGR Is Nothing Then
        For vCount As Integer = 0 To vDGR.Items.Count - 1
          If vDGR.Items(vCount).Cells(mvCPDPointIndex).Text = pContactCPDCycleNumber Then
            vDGR.SelectedIndex = vDGR.Items(vCount).ItemIndex
          End If
        Next
      End If
    Catch vException As Exception
      Throw vException
    End Try
  End Sub

  Private Sub EnableControls(ByVal pEnable As Boolean)
    DirectCast(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).Enabled = pEnable
    DirectCast(Me.FindControl("CpdCategoryType"), DropDownList).Enabled = pEnable
    DirectCast(Me.FindControl("CpdCategory"), DropDownList).Enabled = pEnable
  End Sub

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    'New hyper link should be hidden if the New button is not displayed
    Return If(FindControlByName(mvView2, "New") Is Nothing, False, True)
  End Function
End Class
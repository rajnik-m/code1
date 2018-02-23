Public Class ContactCPDCycle
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Private mvCPDCycleNumber As Integer
  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vCtr As Integer = 0
      InitialiseControls(CareNetServices.WebControlTypes.wctCPDCycle, tblDataEntry)
      If Me.FindControl("PageError") IsNot Nothing Then Me.FindControl("PageError").Visible = False
      If Me.FindControl("WarningMessage") IsNot Nothing Then Me.FindControl("WarningMessage").Visible = False
      If Me.FindControl("StartMonth") IsNot Nothing Then
        DirectCast(Me.FindControl("StartMonth"), TextBox).ReadOnly = False
      End If
      If Me.FindControl("EndMonth") IsNot Nothing Then
        DirectCast(Me.FindControl("EndMonth"), TextBox).ReadOnly = False
      End If
      BindCPDCycleDataGrid()
      HandleMultiViewDisplay()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub BindCPDCycleDataGrid()
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vDataGrid As DataGrid
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      vList("ContactNumber") = UserContactNumber()
      vList("ForPortal") = "Y"
      vDataGrid = TryCast(Me.FindControl("ContactCPDCycle"), DataGrid)
      If vDataGrid IsNot Nothing Then
        Dim vCPDPointPage As String = ""
        Dim vCPDObjectivePage As String = ""
        Dim vHyperlinkText As String = ""
        If InitialParameters.ContainsKey("HyperlinkText") AndAlso InitialParameters("HyperlinkText").ToString.Length > 0 Then
          vHyperlinkText = InitialParameters("HyperlinkText").ToString
        End If
        DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDCycles, vDataGrid, "", UserContactNumber, 0, vList, 0, True, vHyperlinkText)
        If vDataGrid.Columns(1).HeaderText <> "" Then
          Dim vColumn As New BoundColumn()
          vColumn.HeaderText = ""
          vDataGrid.Columns.AddAt(1, vColumn)
          vDataGrid.DataBind()
        End If
        Dim vDataTable As DataTable
        If vDataGrid.Items.Count > 0 Then
          vDataTable = CType(vDataGrid.DataSource, DataSet).Tables("DataRow")
          If InitialParameters.ContainsKey("CPDPointPage") AndAlso InitialParameters("CPDPointPage").ToString.Length > 0 Then
            vCPDPointPage = InitialParameters("CPDPointPage").ToString
          End If
          If InitialParameters.ContainsKey("CPDObjectivePage") AndAlso InitialParameters("CPDObjectivePage").ToString.Length > 0 Then
            vCPDObjectivePage = InitialParameters("CPDObjectivePage").ToString
          End If
          For vRow As Integer = 0 To vDataGrid.Items.Count - 1
            If vDataTable.Rows(vRow).Item("CPDType").ToString = "P" Then
              If vCPDPointPage.Length > 0 Then
                vDataGrid.Items(vRow).Cells(1).Text = "<a href=default.aspx?pn=" & vCPDPointPage & "&CN=" & vDataTable.Rows(vRow).Item("ContactCPDCycleNumber").ToString & ">Details</a>"
              End If
            Else
              If vCPDObjectivePage.Length > 0 Then
                vDataGrid.Items(vRow).Cells(1).Text = "<a href=default.aspx?pn=" & vCPDObjectivePage & "&CN=" & vDataTable.Rows(vRow).Item("ContactCPDCycleNumber").ToString & ">Details</a>"
              End If
            End If
          Next
          If vCPDPointPage.Length = 0 AndAlso vCPDObjectivePage.Length = 0 Then
            vDataGrid.Columns(1).Visible = False
          End If
          If InitialParameters.ContainsKey("AllowEditing") AndAlso InitialParameters("AllowEditing").ToString = "N" Then
            vDataGrid.Columns(0).Visible = False
          End If
        End If

        mvCPDCycleNumber = GetDataGridItemIndex(vDataGrid, "ContactCPDCycleNumber")
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      Throw vException
    End Try
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Dim vDataGrid As DataGrid
      vDataGrid = TryCast(Me.FindControl("ContactCPDCycle"), DataGrid)
      If Me.FindControl("PageError") IsNot Nothing Then Me.FindControl("PageError").Visible = False
      If Me.FindControl("WarningMessage") IsNot Nothing Then Me.FindControl("WarningMessage").Visible = False
      If CType(sender, Control).ID = "Save" AndAlso IsValid() Then
        Dim vList As New ParameterList(HttpContext.Current)
        If IntegerValue(GetTextBoxText("StartYear")) < 1900 Then
          Me.FindControl("WarningMessage").Visible = True
          SetErrorLabel("Start date must not be less than 1900", "WarningMessage")
          If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
          Exit Sub
        End If
        If IntegerValue(GetTextBoxText("EndYear")) > 3000 Then
          Me.FindControl("WarningMessage").Visible = True
          SetErrorLabel("End date must not be greater than 3000", "WarningMessage")
          If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
          Exit Sub
        End If
        'Dim vValid As Boolean = True
        'Dim vUseFlexibleCycles As Boolean = False
        'Dim vDDL As DropDownList = TryCast(FindControlByName(Me, "CpdCycleType"), DropDownList)
        'If vDDL IsNot Nothing Then
        '  Dim vCycleTypeList As New ParameterList(HttpContext.Current)
        '  vCycleTypeList("ForPortal") = "Y"
        '  Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCPDCycleTypes, vCycleTypeList)
        '  Dim vRowColl() As DataRow = vTable.Select("CpdCycleType = '" & vDDL.SelectedValue & "'")
        '  If vRowColl.Length > 0 Then
        '    Dim vRow As DataRow = vRowColl(0)
        '    If vRow IsNot Nothing AndAlso vRow.Item("StartMonth").ToString.Length = 0 Then
        '      vUseFlexibleCycles = True
        '    End If
        '  End If
        'End If

        'Dim vList As New ParameterList(HttpContext.Current)
        'If vUseFlexibleCycles = False Then
        '  If IntegerValue(GetTextBoxText("StartYear")) < 2000 Then
        '    Me.FindControl("PageError").Visible = True
        '    SetErrorLabel("Start year must not be less than 2000", "PageError")
        '    If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        '    Exit Sub
        '  End If
        '  If IntegerValue(GetTextBoxText("EndYear")) < 2000 Then
        '    Me.FindControl("PageError").Visible = True
        '    SetErrorLabel("End year must not be less than 2000", "PageError")
        '    If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        '    Exit Sub
        '  ElseIf IntegerValue(GetTextBoxText("EndYear")) < IntegerValue(GetTextBoxText("StartYear")) Then
        '    Me.FindControl("PageError").Visible = True
        '    SetErrorLabel("End year cannot be less than Start year", "PageError")
        '    If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        '    Exit Sub
        '  End If
        '  vList("StartYear") = GetTextBoxText("StartYear")
        '  vList("EndYear") = GetTextBoxText("EndYear")
        'Else
        '  If HasCPDCycleStartAndEndDates() = False Then
        '    Me.FindControl("PageError").Visible = True
        '    SetErrorLabel("A flexible CPD Cycle Type cannot be used as the Start Date and End Date controls are not visible.", "PageError")
        '    If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        '    Exit Sub
        '  ElseIf GetTextBoxText("StartDate").Length = 0 Then
        '    Me.FindControl("PageError").Visible = True
        '    SetErrorLabel("The start date must be specified for flexible CPD Cycles", "PageError")
        '    If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        '    Exit Sub
        '  ElseIf GetTextBoxText("EndDate").Length = 0 Then
        '    SetCPDEndYearOrDate(False)
        '    If GetTextBoxText("EndDate").Length = 0 Then
        '      Me.FindControl("PageError").Visible = True
        '      SetErrorLabel("The end date must be specified for flexible CPD Cycles", "PageError")
        '      If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        '      Exit Sub
        '    End If
        '  Else
        '    Dim vStartDate As Date
        '    If Date.TryParse(GetTextBoxText("StartDate"), vStartDate) Then
        '      Dim vEndDate As Date = Date.Parse(GetTextBoxText("EndDate"))
        '      Dim vYears As Integer = vEndDate.Year - vStartDate.Year
        '      If vStartDate.Month = 1 Then vYears += 1
        '      If vStartDate.AddYears(vYears).AddDays(-1).CompareTo(vEndDate) <> 0 Then
        '        Me.FindControl("PageError").Visible = True
        '        SetErrorLabel("The CPD Cycle End Date should be a whole number of years less 1 day after the Start Date", "PageError")
        '        If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
        '        Exit Sub
        '      End If
        '    End If
        '  End If
        '  vList("StartDate") = GetTextBoxText("StartDate")
        '  vList("EndDate") = GetTextBoxText("EndDate")
        'End If

        If ViewState("CPDCycleNumber") IsNot Nothing AndAlso ViewState("CPDCycleNumber").ToString.Length > 0 Then
          vList("ContactCpdCycleNumber") = ViewState("CPDCycleNumber").ToString
          vList("ContactNumber") = UserContactNumber()
          vList("CpdCycleStatus") = DirectCast(Me.FindControl("CpdCycleStatus"), DropDownList).SelectedValue
          'vList("CpdCycleType") = DirectCast(Me.FindControl("CpdCycleType"), DropDownList).SelectedValue
          vList("StartYear") = GetTextBoxText("StartYear")
          vList("EndYear") = GetTextBoxText("EndYear")
          DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctCPDCycles, vList)
          'ProcessRedirect(String.Format("Default.aspx?pn={0}&CN={1}", Request.QueryString("pn").ToString, Request.QueryString("CN").ToString))
        Else
          vList("ContactNumber") = UserContactNumber()
          vList("CpdCycleStatus") = DirectCast(Me.FindControl("CpdCycleStatus"), DropDownList).SelectedValue
          vList("CpdCycleType") = DirectCast(Me.FindControl("CpdCycleType"), DropDownList).SelectedValue
          vList("StartYear") = GetTextBoxText("StartYear")
          vList("EndYear") = GetTextBoxText("EndYear")
          DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctCPDCycles, vList)
          If vDataGrid IsNot Nothing AndAlso vDataGrid.Items.Count = 0 Then
            ProcessRedirect(String.Format("Default.aspx?pn={0}", Request.QueryString("pn").ToString))
          End If
        End If
        BindCPDCycleDataGrid()
        HighLightDataGridRow()
      ElseIf (CType(sender, Control).ID = "New" OrElse CType(sender, Control).ID = "GridHyperlink") And Not InWebPageDesigner() Then
        SetTextBoxText("StartMonth", "")
        SetTextBoxText("StartYear", "")
        SetTextBoxText("EndMonth", "")
        SetTextBoxText("EndYear", "")
        'If HasCPDCycleStartAndEndDates() Then
        '  SetTextBoxText("StartDate", "")
        '  SetTextBoxText("EndDate", "")
        'End If
        ViewState("CPDCycleNumber") = ""
        DirectCast(Me.FindControl("CpdCycleType"), DropDownList).Enabled = True
        DirectCast(Me.FindControl("CpdCycleType"), DropDownList).SelectedIndex = 0
        DirectCast(Me.FindControl("CpdCycleStatus"), DropDownList).SelectedIndex = 0
        SetControlEnabled("EndYear", True)
        SetControlEnabled("StartMonth", True)
        SetControlEnabled("EndMonth", True)
        'SetControlEnabled("StartYear", True)
        If Me.FindControl("PageError") IsNot Nothing Then SetLabelText("PageError", "")
        If Me.FindControl("WarningMessage") IsNot Nothing Then SetLabelText("WarningMessage", "")
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      Me.FindControl("PageError").Visible = True
      SetLabelText("PageError", vException.Message)
      If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
    End Try
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Try
      Dim vCPDCycleNumber As String = DirectCast(e.Item.Cells(mvCPDCycleNumber).Controls(0), ITextControl).Text
      ViewState("CPDCycleNumber") = vCPDCycleNumber

      Dim vList As New ParameterList(HttpContext.Current)
      'Get Contact Position Details
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      vList("ContactNumber") = UserContactNumber()
      vList("ForPortal") = "Y"
      vList("ContactCpdCycleNumber") = vCPDCycleNumber
      Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDCycles, vList))
      If vRow IsNot Nothing Then
        Dim vCycleList As ParameterList
        Dim vDDL As DropDownList
        DirectCast(Me.FindControl("CpdCycleType"), DropDownList).SelectedValue = vRow.Item("CpdCycleType").ToString
        'If vRow.Item("StartMonth").ToString.Length = 0 Then
        '  'Flexible Cycle
        '  SetTextBoxText("StartDate", vRow.Item("StartDate").ToString)
        '  SetTextBoxText("EndDate", vRow.Item("EndDate").ToString)
        '  SetControlEnabled("StartYear", False)
        'Else
        '  'Fixed Cycle
        SetTextBoxText("StartMonth", MonthName(IntegerValue(vRow.Item("StartMonth").ToString)))
        SetTextBoxText("StartYear", Substring(vRow.Item("CycleStart").ToString, 4, 4))
        SetTextBoxText("EndMonth", MonthName(IntegerValue(vRow.Item("EndMonth").ToString)))
        SetTextBoxText("EndYear", Substring(vRow.Item("CycleEnd").ToString, 4, 4))
        'End If
        vCycleList = New ParameterList(HttpContext.Current)
        vDDL = DirectCast(Me.FindControl("CpdCycleStatus"), DropDownList)
        vCycleList("CPDCycleType") = vRow.Item("CpdCycleType").ToString
        vCycleList("ForPortal") = "Y"
        vDDL.DataTextField = "CPDCycleStatusDesc"
        vDDL.DataValueField = "CPDCycleStatus"
        DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDCycleStatuses, vDDL, True, vCycleList)
        DirectCast(Me.FindControl("CpdCycleStatus"), DropDownList).SelectedValue = vRow.Item("CpdCycleStatus").ToString
        DirectCast(Me.FindControl("CpdCycleType"), DropDownList).Enabled = False
        SetControlEnabled("StartMonth", False)
        SetControlEnabled("EndMonth", False)
        SetControlEnabled("EndYear", False)
      End If

    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Private Sub HighLightDataGridRow()
    Try
      Dim vDGR As DataGrid = TryCast(Me.FindControl("ContactCPDCycle"), DataGrid)
      If Not vDGR Is Nothing Then
        If ViewState("CPDCycleNumber") IsNot Nothing Then
          For vCount As Integer = 0 To vDGR.Items.Count - 1
            If vDGR.Items(vCount).Cells(mvCPDCycleNumber).Text = ViewState("CPDCycleNumber").ToString Then
              vDGR.SelectedIndex = vDGR.Items(vCount).ItemIndex
              vDGR.SelectedIndex = vDGR.Items(vCount).ItemIndex
            End If
          Next
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      Throw vException
    End Try
  End Sub

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    'New hyper link should be hidden if the New button is not displayed
    Return If(FindControlByName(mvView2, "New") Is Nothing, False, True)
  End Function

End Class
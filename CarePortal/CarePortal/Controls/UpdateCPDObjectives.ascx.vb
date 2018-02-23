Public Class UpdateCPDObjectives
  Inherits CareWebControl
  Implements IMultiViewWebControl

  Private mvCPDObjectiveIndex As Integer

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try

      Dim vList As New ParameterList(HttpContext.Current)

      If Not InWebPageDesigner() AndAlso Request.QueryString("CN") Is Nothing Then Throw New Exception("Parameter CycleNumber is missing")
      If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then
        vList("ContactCpdCycleNumber") = IntegerValue(Request.QueryString("CN").ToString)
      End If

      Dim vCtr As Integer = 0
      InitialiseControls(CareNetServices.WebControlTypes.wctUpdateCpdObjectives, tblDataEntry)
      If TryCast(Me.FindControl("ContactCPDObjectives"), DataGrid) IsNot Nothing Then
        Me.FindControl("PageError").Visible = False
        vList("SystemColumns") = "Y"
        vList("ContactNumber") = UserContactNumber()
        vList("ForPortal") = "Y"

        BindCPDObjectiveDataGrid()

        'Populate Period Combo-Box
        Dim vCPDPeriod As DropDownList = CType(Me.FindControl("ContactCpdPeriodNumber"), DropDownList)
        vCPDPeriod.DataTextField = "ContactCpdPeriodNumberDesc"
        vCPDPeriod.DataValueField = "ContactCpdPeriodNumber"
        DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtCPDPeriods, vCPDPeriod, True, vList)
      End If
      HandleMultiViewDisplay()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub
  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If CType(sender, Control).ID = "Save" AndAlso IsValid() Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("UserID") = UserContactNumber.ToString
        'If 'Save' is clicked for Edit operation
        If ViewState("CPDObjectiveNumber") IsNot Nothing Then
          vList("CompletionDate") = GetTextBoxText("CompletionDate")
          vList("ContactCpdPeriodNumber") = GetDropDownValue("ContactCpdPeriodNumber")
          vList("ContactNumber") = UserContactNumber()
          vList("CpdCategory") = GetDropDownValue("CpdCategory")
          vList("CpdCategoryType") = GetDropDownValue("CpdCategoryType")
          vList("CpdObjectiveDesc") = GetTextBoxText("CpdObjectiveDesc")
          vList("CpdObjectiveNumber") = ViewState("CPDObjectiveNumber")
          vList("LongDescription") = GetTextBoxText("LongDescription")
          vList("Notes") = GetTextBoxText("Notes")
          vList("TargetDate") = GetTextBoxText("TargetDate")

          DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctCPDObjectives, vList)
          BindCPDObjectiveDataGrid()
          HighLightDataGridRow(ViewState("CPDObjectiveNumber").ToString)
        Else
          'If 'Save' is clicked for New operation
          vList("CompletionDate") = GetTextBoxText("CompletionDate")
          vList("ContactCpdPeriodNumber") = GetDropDownValue("ContactCpdPeriodNumber")
          vList("ContactNumber") = UserContactNumber()
          vList("CpdCategory") = GetDropDownValue("CpdCategory")
          vList("CpdCategoryType") = GetDropDownValue("CpdCategoryType")
          vList("CpdObjectiveDesc") = GetTextBoxText("CpdObjectiveDesc")
          vList("LongDescription") = GetTextBoxText("LongDescription")
          vList("Notes") = GetTextBoxText("Notes")
          vList("TargetDate") = GetTextBoxText("TargetDate")
          vList("SupervisorAccepted") = "N"

          Dim vParamList As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctCPDObjectives, vList)
          BindCPDObjectiveDataGrid()
          HighLightDataGridRow(vParamList.Item("CpdObjectiveNumber").ToString)
          ViewState("CPDObjectiveNumber") = vParamList.Item("CpdObjectiveNumber").ToString
        End If
        CType(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).Enabled = False
        CType(Me.FindControl("CpdCategoryType"), DropDownList).Enabled = False
        CType(Me.FindControl("CpdCategory"), DropDownList).Enabled = False
      ElseIf (CType(sender, Control).ID = "New" OrElse CType(sender, Control).ID = "GridHyperlink") AndAlso Not InWebPageDesigner() Then
        SetDropDownText("ContactCpdPeriodNumber", String.Empty)
        SetDropDownText("CpdCategoryType", String.Empty, True)
        CType(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).Enabled = True
        CType(Me.FindControl("CpdCategoryType"), DropDownList).Enabled = True
        CType(Me.FindControl("CpdCategory"), DropDownList).Enabled = True

        SetTextBoxText("CpdObjectiveDesc", String.Empty)
        SetTextBoxText("LongDescription", String.Empty)
        SetTextBoxText("CompletionDate", String.Empty)
        SetTextBoxText("TargetDate", String.Empty)
        SetTextBoxText("Notes", String.Empty)

        ViewState.Remove("CPDObjectiveNumber")
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      Me.FindControl("PageError").Visible = True
      SetLabelText("PageError", vException.Message)
      If mvSupportsMultiView Then mvMultiView.SetActiveView(mvView2)
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Try
      Dim vCPDObjectiveNumber As String = DirectCast(e.Item.Cells(mvCPDObjectiveIndex).Controls(0), ITextControl).Text
      ViewState("CPDObjectiveNumber") = vCPDObjectiveNumber

      Dim vList As New ParameterList(HttpContext.Current)
      'Get Contact Position Details
      vList("ContactNumber") = UserContactNumber()
      vList("CpdObjectiveNumber") = vCPDObjectiveNumber
      Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDObjectives, vList))
      If vRow IsNot Nothing Then
        CType(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).SelectedValue = vRow("ContactCpdPeriodNumber").ToString
        CType(Me.FindControl("ContactCpdPeriodNumber"), DropDownList).Enabled = False
        SetDropDownText("CpdCategoryType", vRow("CategoryType").ToString, True)
        CType(Me.FindControl("CpdCategoryType"), DropDownList).Enabled = False
        CType(Me.FindControl("CpdCategory"), DropDownList).SelectedValue = vRow("Category").ToString
        CType(Me.FindControl("CpdCategory"), DropDownList).Enabled = False
        SetTextBoxText("CpdObjectiveDesc", vRow("CpdObjectiveDesc").ToString)
        SetTextBoxText("LongDescription", vRow("LongDescription").ToString)
        SetTextBoxText("CompletionDate", vRow("CompletionDate").ToString)
        SetTextBoxText("TargetDate", vRow("TargetDate").ToString)
        SetTextBoxText("Notes", vRow("Notes").ToString)
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub HighLightDataGridRow(ByVal pCPDObjectiveIndex As String)
    Try
      Dim vDGR As DataGrid = CType(Me.FindControl("ContactCPDObjectives"), DataGrid)
      If Not vDGR Is Nothing Then
        For vCount As Integer = 0 To vDGR.Items.Count - 1
          If vDGR.Items(vCount).Cells(mvCPDObjectiveIndex).Text = pCPDObjectiveIndex Then
            vDGR.SelectedIndex = vDGR.Items(vCount).ItemIndex
          End If
        Next
      End If
    Catch vException As Exception
      Throw vException
    End Try
  End Sub

  Private Sub BindCPDObjectiveDataGrid()
    Dim vDGR As DataGrid = CType(Me.FindControl("ContactCPDObjectives"), DataGrid)
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vHyperlinkText As String = ""
    vList("SystemColumns") = "Y"
    vList("ContactNumber") = UserContactNumber()
    vList("ForPortal") = "Y"
    vList("WebPageItemNumber") = Me.WebPageItemNumber
    If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then
      vList("ContactCpdCycleNumber") = IntegerValue(Request.QueryString("CN").ToString)
    End If
    vDGR.Columns.Clear()
    If InWebPageDesigner() Then
      vList("DocumentColumns") = "Y"
    End If
    If InitialParameters.ContainsKey("HyperlinkText") AndAlso InitialParameters("HyperlinkText").ToString.Length > 0 Then
      vHyperlinkText = InitialParameters("HyperlinkText").ToString
    End If
    DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCPDObjectives, vDGR, String.Empty, UserContactNumber, , vList, 0, True, vHyperlinkText)
    mvCPDObjectiveIndex = GetDataGridItemIndex(vDGR, "CpdObjectiveNumber")
    If InitialParameters.ContainsKey("AllowEditing") AndAlso InitialParameters("AllowEditing").ToString = "N" Then
      vDGR.Columns(0).Visible = False
    End If
  End Sub

  Public Function GridHyperLinkVisibility() As Boolean Implements IMultiViewWebControl.GridHyperLinkVisibility
    'New hyper link should be hidden if the New button is not displayed
    Return If(FindControlByName(mvView2, "New") Is Nothing, False, True)
  End Function
End Class
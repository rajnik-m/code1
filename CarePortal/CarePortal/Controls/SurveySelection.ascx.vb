Public Class SurveySelection
  Inherits CareWebControl
  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      Dim vList As New ParameterList(HttpContext.Current)
      Dim vDataGrid As DataGrid
      InitialiseControls(CareNetServices.WebControlTypes.wctSelectSurveys, tblDataEntry)
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      vList("ContactNumber") = UserContactNumber().ToString
      If InitialParameters.ContainsKey("SurveyType") AndAlso InitialParameters("SurveyType").ToString.Length > 0 Then
        vList("SurveyType") = InitialParameters("SurveyType").ToString
      End If
      If InitialParameters.ContainsKey("RegisteredSurveyType") AndAlso InitialParameters("RegisteredSurveyType").ToString.Length > 0 Then
        vList("RegisteredSurveyType") = InitialParameters("RegisteredSurveyType").ToString
      End If
      vDataGrid = TryCast(Me.FindControl("SurveyData"), DataGrid)
      If vDataGrid IsNot Nothing Then
        Dim pList As New ParameterList(HttpContext.Current)
        Dim vContactResult As String = String.Empty
        Dim vDataTable As DataTable
        Dim vResult As String = String.Empty
        Dim vSurveyEntryPage As String = String.Empty
        Dim vColumn As New BoundColumn()
        Dim vSurveyVersionPos As Integer

        pList("ContactNumber") = UserContactNumber().ToString
        pList("SystemColumns") = "Y"
        vContactResult = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactHeaderInformation, pList)
        vDataTable = GetDataTable(vContactResult, True)
        vList("ContactGroup") = vDataTable.Rows(0).Item("GroupCode").ToString

        vResult = DataHelper.FindData(CareNetServices.XMLDataFinderTypes.xdftWebSurveys, vList)
        DataHelper.FillGrid(vResult, vDataGrid)

        If InitialParameters.Contains("SurveyEntryPage") Then vSurveyEntryPage = InitialParameters("SurveyEntryPage").ToString

        vColumn.HeaderText = ""
        vDataGrid.Columns.AddAt(0, vColumn)
        vDataGrid.DataBind()
        For vCount As Integer = 0 To vDataGrid.Columns.Count - 1
          Dim vBoundColumn As BoundColumn = DirectCast(vDataGrid.Columns(vCount), BoundColumn)
          If vBoundColumn.DataField = "SurveyVersionNumber" Then
            vSurveyVersionPos = vCount
          End If
        Next
        If vSurveyEntryPage.Length = 0 Then
          vDataGrid.Columns(0).Visible = False
        Else
          For vRow As Integer = 0 To vDataGrid.Items.Count - 1
            vDataGrid.Items(vRow).Cells(0).Text = "<a href='default.aspx?pn=" & vSurveyEntryPage & "&SV=" & vDataGrid.Items(vRow).Cells(vSurveyVersionPos).Text & "'>Select</a>"
          Next
        End If
        If vDataGrid.Items.Count <= 0 Then
          vDataGrid.Visible = False
        Else
          DirectCast(Me.FindControl("WarningMessage"), Label).Visible = False
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    End Try
  End Sub

End Class
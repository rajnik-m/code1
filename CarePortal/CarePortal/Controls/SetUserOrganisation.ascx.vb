Public Class SetUserOrganisation
  Inherits CareWebControl

  Dim mvWarningMessage As New Label
  Dim mvOrganisationNumberIndex As Integer = 3

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctSetUserOrganisation, tblDataEntry)
      mvWarningMessage = CType(FindControlByName(Me, "WarningMessage1"), Label)
      mvWarningMessage.Visible = False
      ShowOrganisations()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As PortalAccessException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  ''' <summary>
  ''' Show available organisations to allow user selection
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub ShowOrganisations()
    Dim vList As New ParameterList(HttpContext.Current)
    vList("ViewType") = "U"

    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtAllViews, vList)
    Dim vResult As String
    vList = New ParameterList(HttpContext.Current)
    vList("ViewName") = vDataTable.Rows(0)("ViewName")
    vList("Filter") = String.Format("contact_number = {0}", UserContactNumber())
    vList("SelectedColumns") = "position,organisation_name,organisation_number"
    vList("SelectedHeadings") = "Position,Organisation Name,Organisation Number"
    vResult = DataHelper.SelectTableDataString(CareNetServices.XMLTableDataSelectionTypes.xtdstViewData, vList)

    Dim vOrganisationData As DataTable = GetDataTable(vResult, True)
    vOrganisationData = GetDataTable(vResult, True)

    Dim vOrgDataGrid As DataGrid
    vOrgDataGrid = CType(FindControlByName(Me, "UserOrganisationData"), DataGrid)

    If vOrganisationData Is Nothing OrElse vOrganisationData.Rows.Count = 0 Then
      mvWarningMessage.Visible = True
    ElseIf vOrganisationData.Rows.Count = 1 Then
      Session("SelectedOrganisationNumber") = vOrganisationData.Rows(0).Item("Organisation_Number")
      GoToSubmitPage()
    Else
      FindControlByName(Me, "UserOrganisationData").Visible = True
      DataHelper.FillGrid(vResult, vOrgDataGrid, , , pDisplayEditColumn:=True, pCommandNameForEditColumn:="Select")
      SetControlVisible("UserOrganisationData", True)
    End If
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    SetOrganisation(e.Item)
  End Sub

  Private Sub SetOrganisation(ByVal pDataGridItem As DataGridItem)
    If DirectCast(pDataGridItem.Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text.Length > 0 Then
      Session("SelectedOrganisationNumber") = DirectCast(pDataGridItem.Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text
      GoToSubmitPage()
    End If
  End Sub

End Class
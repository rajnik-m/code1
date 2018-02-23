Public Class AddPosition
  Inherits CareWebControl
  Private mvAddressNumber As Integer
  Public Sub New()
    mvNeedsAuthentication = True
  End Sub
  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddPosition, tblDataEntry)
      If Not InWebPageDesigner() Then
        CheckSessionValueSet()
        SetDefaults()
      End If
      SetControlEnabled("Name", False)
      SetControlEnabled("Address", False)
      SetErrorLabel("")
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  Private Sub SetDefaults()
    Dim vOrganisationNumber As Integer = IntegerValue(Session("SelectedOrganisationNumber").ToString)
    Dim vContactNumber As Integer = IntegerValue(Session("SelectedContactNumber").ToString)
    Dim vList As New ParameterList(HttpContext.Current)
    vList("ContactNumber") = vOrganisationNumber

    Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vList))
    If vRow IsNot Nothing Then
      SetTextBoxText("Name", vRow("ContactName").ToString)
      SetTextBoxText("Address", vRow("AddressLine").ToString)
      If InitialParameters("Position") IsNot Nothing Then
        SetTextBoxText("Position", InitialParameters("Position").ToString)
      Else
        SetTextBoxText("Position", "")
      End If
      If InitialParameters("PositionFunction") IsNot Nothing Then
        SetDropDownText("PositionFunction", InitialParameters("PositionFunction").ToString)
      Else
        SetDropDownText("PositionFunction", "")
      End If
      If InitialParameters("PositionSeniority") IsNot Nothing Then
        SetDropDownText("PositionSeniority", InitialParameters("PositionSeniority").ToString)
      Else
        SetDropDownText("PositionSeniority", "")
      End If
      mvAddressNumber = IntegerValue(vRow("AddressNumber").ToString)
    End If
  End Sub
  Private Sub CheckSessionValueSet()
    If Session("SelectedContactNumber") Is Nothing OrElse Session("SelectedContactNumber").ToString.Length = 0 _
      OrElse Session("SelectedOrganisationNumber") Is Nothing AndAlso Session("SelectedOrganisationNumber").ToString.Length = 0 Then
      Throw New PortalAccessException
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If Not InWebPageDesigner() Then
        If IsValid() Then
          Dim vOrganisationNumber As Integer = IntegerValue(Session("SelectedOrganisationNumber").ToString)
          Dim vContactNumber As Integer = IntegerValue(Session("SelectedContactNumber").ToString)
          Dim vList As New ParameterList(HttpContext.Current)
          vList("ContactNumber") = vContactNumber
          vList("OrganisationNumber") = vOrganisationNumber
          vList("AddressNumber") = mvAddressNumber
          If GetTextBoxText("PositionLocation").Length > 0 Then vList("Location") = GetTextBoxText("PositionLocation")

          'Set vList param to Default value if the control is invisible
          'Position
          If FindControl("Position") IsNot Nothing AndAlso GetTextBoxText("Position").Length > 0 Then
            vList("Position") = GetTextBoxText("Position")
          Else
            If InitialParameters("Position") IsNot Nothing Then
              vList("Position") = InitialParameters("Position").ToString
            End If
          End If
          'Function
          If FindControl("PositionFunction") IsNot Nothing AndAlso GetDropDownValue("PositionFunction").Length > 0 Then
            vList("PositionFunction") = GetDropDownValue("PositionFunction")
          Else
            If InitialParameters("PositionFunction") IsNot Nothing Then
              vList("PositionFunction") = InitialParameters("PositionFunction").ToString
            End If
          End If
          'Seniority
          If FindControl("PositionSeniority") IsNot Nothing AndAlso GetDropDownValue("PositionSeniority").Length > 0 Then
            vList("PositionSeniority") = GetDropDownValue("PositionSeniority")
          Else
            If InitialParameters("PositionSeniority") IsNot Nothing Then
              vList("PositionSeniority") = InitialParameters("PositionSeniority").ToString
            End If
          End If

          If GetTextBoxText("Started").Length > 0 Then vList("ValidFrom") = GetTextBoxText("Started")
          If GetTextBoxText("Finished").Length > 0 Then vList("ValidTo") = GetTextBoxText("Finished")
          vList("Mail") = BooleanString(GetCheckBoxChecked("Mail"))
          vList("AdjustNullDates") = "Y"

          DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctPosition, vList)
          GoToSubmitPage()
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enPositionDatesExceedSiteDates Then
        SetErrorLabel(vEx.Message)
      Else
        ProcessError(vEx)
      End If

    End Try
  End Sub

End Class
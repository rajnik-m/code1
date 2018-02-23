Public Class PayerSelection
  Inherits CareWebControl

  Private mvOrganisationNumberIndex As Integer = -1
  Private mvAddressNumberIndex As Integer = -1

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctPayerSelection, tblDataEntry, "", "")
      Dim vPayerDataGrid As DataGrid = CType(FindControlByName(Me, "PayerData"), DataGrid)
      vPayerDataGrid.Columns.Clear()
      Dim vList As New ParameterList(HttpContext.Current)
      vList("SystemColumns") = "Y"
      vList("WebPageItemNumber") = Me.WebPageItemNumber
      DataHelper.GetContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddressesAndPositions, vPayerDataGrid, "", UserContactNumber, 0, vList, 0, True, "Select")
      mvOrganisationNumberIndex = GetDataGridItemIndex(vPayerDataGrid, "OrganisationNumber")
      mvAddressNumberIndex = GetDataGridItemIndex(vPayerDataGrid, "AddressNumber")

      'Navigate to Submit Page when 
      'A. There is no current address
      'B. There is only one current private address and no position
      'C. There is only one current address which is for a position and the BypassIfOneAddress parameter is set 
      If Not InWebPageDesigner() Then
        If vPayerDataGrid.Items.Count = 0 Then
          'No current address. Just navigate to the Submit page. (ProcessPayment will use the user default address)
          Session.Remove("PayerContactNumber")
          Session.Remove("PayerAddressNumber")
          GoToSubmitPage()
        ElseIf vPayerDataGrid.Items.Count = 1 Then
          'There is only one current address. Do not display it to the user and set the payer info from this record.
          SetPayerInfo(vPayerDataGrid.Items(0))
        ElseIf vPayerDataGrid.Items.Count = 2 AndAlso InitialParameters.OptionalValue("BypassIfOneAddress") = "Y" Then
          If DirectCast(vPayerDataGrid.Items(0).Cells(mvAddressNumberIndex).Controls(0), ITextControl).Text.Equals(DirectCast(vPayerDataGrid.Items(1).Cells(mvAddressNumberIndex).Controls(0), ITextControl).Text) Then
            'Both rows are for the same address
            Dim vOrg0 As String = DirectCast(vPayerDataGrid.Items(0).Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text
            If vOrg0 = "&nbsp;" Then vOrg0 = ""
            Dim vOrg1 As String = DirectCast(vPayerDataGrid.Items(1).Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text
            If vOrg1 = "&nbsp;" Then vOrg1 = ""
            If vOrg0.Length > 0 Then
              SetPayerInfo(vPayerDataGrid.Items(0))
            Else
              SetPayerInfo(vPayerDataGrid.Items(1))
            End If
          End If
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    'On selecting a row, just set the payer number and address number in the session and go to Submit page.
    SetPayerInfo(e.Item)
  End Sub

  Private Sub SetPayerInfo(ByVal pDataGridItem As DataGridItem)
    If DirectCast(pDataGridItem.Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text <> "&nbsp;" AndAlso DirectCast(pDataGridItem.Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text.Length > 0 Then
      Session("PayerContactNumber") = DirectCast(pDataGridItem.Cells(mvOrganisationNumberIndex).Controls(0), ITextControl).Text
    Else
      Session("PayerContactNumber") = UserContactNumber()
    End If
    Session("PayerAddressNumber") = DirectCast(pDataGridItem.Cells(mvAddressNumberIndex).Controls(0), ITextControl).Text

    GoToSubmitPage()
  End Sub

End Class
Public Class DeDupOrgForRegistration
  Inherits CareWebControl

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctDeDupOrgForRegistration, tblDataEntry, "", "")
      'If WarningMessage1 is set visible in module customisation then hide it.
      If FindControlByName(Me, "WarningMessage1") IsNot Nothing Then FindControlByName(Me, "WarningMessage1").Visible = False

      'Only try populating the grid if we are coming from Register module
      If Session("AddContactList") IsNot Nothing Then
        Dim vDuplicateList As New ParameterList(HttpContext.Current)
        vDuplicateList("Name") = DirectCast(Session("AddContactList"), ParameterList)("Name")
        vDuplicateList("SystemColumns") = "Y"
        vDuplicateList("WebPageItemNumber") = Me.WebPageItemNumber
        Dim vResult As String = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtDuplicateOrganisationsForRegistration, vDuplicateList)
        Dim vOrganisationData As DataTable = GetDataTable(vResult, True)

        'Check for MaximumOrganisations we can display. The default is 20.
        Dim vMaxOrganisations As Integer = If(InitialParameters("MaximumOrganisations").ToString.Length > 0, IntegerValue(InitialParameters("MaximumOrganisations").ToString), 20)

        If IsPostBack = False AndAlso vOrganisationData IsNot Nothing AndAlso vOrganisationData.Rows.Count > vMaxOrganisations Then
          'If the WarningMessage1 was set visible in module customisation then make it visible.
          If FindControlByName(Me, "WarningMessage1") IsNot Nothing Then FindControlByName(Me, "WarningMessage1").Visible = True
        Else
          Dim vDuplicateOrganisations As DataGrid = CType(FindControlByName(Me, "DuplicateOrganisations"), DataGrid)
          vDuplicateOrganisations.Columns.Clear()
          DataHelper.FillGrid(vResult, vDuplicateOrganisations, pDisplayEditColumn:=True, pCommandNameForEditColumn:="Select")
        End If
      End If
      If Request.QueryString("RegURL") Is Nothing Then
        'If we are not coming from Register module then just add the java script to navigate to the previous page.
        Dim vBackButton As Button = TryCast(FindControlByName(Me, "Back"), Button)
        If vBackButton IsNot Nothing Then vBackButton.Attributes.Add("onClick", "javascript:history.back(); return false;")
      End If

    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If Request.QueryString("RegURL") IsNot Nothing Then
      'If RegURL exists then navigate back to the Register module otherwise the java script will be used.
      ProcessRedirect(Request.QueryString("RegURL").ToString)
    End If
  End Sub

  Public Overrides Sub HandleDataGridEdit(ByVal e As DataGridCommandEventArgs)
    Try
      If Session("AddContactList") IsNot Nothing Then 'Make sure the session is still valid
        Dim vOrganisationNumber As String = String.Empty
        Dim vAddressNumber As String = String.Empty
        'Get the Organisation & Address numbers for the currently selected Organisation
        Dim vDGR As DataGrid = CType(FindControlByName(Me, "DuplicateOrganisations"), DataGrid)
        Dim vDT As DataTable = Nothing
        Dim vDR As DataRow = Nothing
        If TypeOf (vDGR.DataSource) Is DataSet AndAlso CType(vDGR.DataSource, DataSet).Tables.Contains("DataRow") Then
          vDT = CType(vDGR.DataSource, DataSet).Tables("DataRow")
        End If
        If vDT IsNot Nothing AndAlso vDT.Rows.Count >= e.Item.ItemIndex Then
          vDR = vDT.Rows(e.Item.ItemIndex)  'This is the row selected by the user
        End If
        If vDR IsNot Nothing Then
          vOrganisationNumber = vDR.Item("OrganisationNumber").ToString
          vAddressNumber = vDR.Item("AddressNumber").ToString
          'Debug.WriteLine("Organisation Name:" & vDR.Item("OrganisationName").ToString)
          'Debug.WriteLine("Address:" & vDR.Item("AddressLine").ToString)
        End If

        If DefaultParameters.ContainsKey("ReturnToRegisterPage") AndAlso DefaultParameters("ReturnToRegisterPage").ToString = "Y" AndAlso Request.QueryString("RegURL") IsNot Nothing Then
          ProcessRedirect(Request.QueryString("RegURL").ToString & String.Format("&ON={0}&AN={1}", vOrganisationNumber, vAddressNumber))
        Else
          Dim vAddContactList As ParameterList = DirectCast(Session("AddContactList"), ParameterList)
          vAddContactList("AddressNumber") = vAddressNumber
          vAddContactList("CheckAdditionalAddress") = "Y" 'This will allow saving a personal address to the contact as well if address fields are provided
          If vOrganisationNumber.Length > 0 Then vAddContactList("OrganisationNumber") = vOrganisationNumber 'pass org number through so we know we selected existing org  for de duping contact
          If vAddContactList.ContainsKey("ConfirmEMailAddress") Then vAddContactList.Remove("ConfirmEMailAddress")
          Dim vList As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vAddContactList)
          Session.Remove("AddContactList")
          If Session("EmailParams") IsNot Nothing Then  'Confirmation is required
            DataHelper.ProcessBulkEMail(DirectCast(Session("ContentParams"), ParameterList).ToCSVFile, DirectCast(Session("EmailParams"), ParameterList), True)
            Session.Remove("ContentParams")
            Session.Remove("EmailParams")
          Else
            Session("ContactNumber") = vList("ContactNumber")
            Session("AddressNumber") = vList("AddressNumber")
            SetAuthentication(vList, True)
          End If
          GoToSubmitPage()
        End If
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

End Class
Public Class UpdateOrganisation
  Inherits CareWebControl

  Private mvOrgNum As String
  Private mvAddNum As String

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctUpdateOrganisation, tblDataEntry, String.Empty, "SwitchboardNumber,WebAddress,FaxNumber")
      CheckAccessRights()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vEx As PortalAccessOrganisationException
      Throw vEx
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  ''' <summary>
  ''' Checks if the user has rights to update the organisations details
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CheckAccessRights()
    Dim vList As New ParameterList(HttpContext.Current)
    vList("ViewType") = "U"
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtAllViews, vList)
    If Not InWebPageDesigner() AndAlso vDataTable IsNot Nothing OrElse (Session("SelectedOrganisationNumber") IsNot Nothing AndAlso Session("SelectedOrganisationNumber").ToString.Length > 0) Then
      If Session("SelectedOrganisationNumber") IsNot Nothing AndAlso Session("SelectedOrganisationNumber").ToString.Length > 0 Then
        'If Session vlaue is set, User already has access to the Organisation specified in session variable. 
        'No need to check for access rights.
        vList = New ParameterList(HttpContext.Current)
        'Get Organisation Details
        vList("ContactNumber") = Session("SelectedOrganisationNumber").ToString
        vDataTable = GetDataTable(DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactInformation, vList))
        With vDataTable.Rows(0)
          SetTextBoxText("Name", .Item("OrganisationName").ToString)
          SetDropDownText("Country", .Item("CountryCode").ToString)
          SetTextBoxText("Postcode", .Item("Postcode").ToString)
          SetTextBoxText("BuildingNumber", .Item("BuildingNumber").ToString)
          SetTextBoxText("HouseName", .Item("HouseName").ToString)
          SetTextBoxText("Address", .Item("Address").ToString)
          SetTextBoxText("Town", .Item("Town").ToString)
          SetTextBoxText("County", .Item("County").ToString)

          'Get contact comm info to populate remaining fields
          vList = New ParameterList(HttpContext.Current)

          mvOrgNum = .Item("OrganisationNumber").ToString
          mvAddNum = .Item("AddressNumber").ToString

          vList("ContactNumber") = .Item("OrganisationNumber")
          vList("AddressNumber") = .Item("AddressNumber")
        End With
      Else
        'Check if the user has access to the first view if there are multiple rows that are returned
        vList = New ParameterList(HttpContext.Current)
        vList("ViewName") = vDataTable.Rows(0)("ViewName")
        vList("Filter") = String.Format("contact_number = {0}", GetContactNumberFromParentGroup)
        vDataTable = DataHelper.SelectTableData(CareNetServices.XMLTableDataSelectionTypes.xtdstViewData, vList)
        If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
          'Populate the organisation name and the address using the first returned row
          With vDataTable.Rows(0)
            SetTextBoxText("Name", .Item("organisation_name").ToString)
            SetDropDownText("Country", .Item("country").ToString)
            SetTextBoxText("Postcode", .Item("postcode").ToString)
            SetTextBoxText("BuildingNumber", .Item("building_number").ToString)
            SetTextBoxText("HouseName", .Item("house_name").ToString)
            SetTextBoxText("Address", .Item("address").ToString)
            SetTextBoxText("Town", .Item("town").ToString)
            SetTextBoxText("County", .Item("county").ToString)

            'Get contact comm info to populate remaining fields
            vList = New ParameterList(HttpContext.Current)

            mvOrgNum = .Item("organisation_number").ToString
            mvAddNum = .Item("address_number").ToString

            vList("ContactNumber") = .Item("organisation_number")
            vList("AddressNumber") = .Item("address_number")
          End With
        Else
          Throw New PortalAccessOrganisationException()
        End If
      End If
      'Get contact comm info to populate remaining fields
      Dim vContactComm As DataTable = GetDataTable(DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsInformation, vList))
      If vContactComm IsNot Nothing AndAlso vContactComm.Rows.Count > 0 Then
        SetTextBoxText("WebAddress", vContactComm.Rows(0).Item("WebAddress").ToString)
        SetTextBoxText("SwitchboardNumber", vContactComm.Rows(0).Item("SwitchBoardNumber").ToString)
        SetTextBoxText("FaxNumber", vContactComm.Rows(0).Item("FaxNumber").ToString)
      End If
    End If
  End Sub

  Public Overrides Sub ProcessSubmit()
    Dim vList As New ParameterList(HttpContext.Current)
    vList = GetAddOrganisationParameterList()
    'Add additional params
    vList("ContactNumber") = mvOrgNum
    vList("AddressNumber") = mvAddNum
    vList("UserID") = UserContactNumber.ToString
    DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vList)
  End Sub

End Class
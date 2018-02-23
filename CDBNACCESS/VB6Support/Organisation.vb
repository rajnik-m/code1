Namespace Access

  Public Enum OrganisationRecordSetTypes 'These are bit values
    ortNumber = 1
    ortName = 2
    ortGroup = 4
    ortPhone = 8
    ortDetail = 16
    ortAll = 31 'Include Number, Name and all other Org table details
    ortAllFields = 31 'All organisation fields only
    ortAddress = 32
    ortAddressCountry = 64
    ortDefaultContact = 128
    ortAllWithDetails = &HFFFFS 'All Org, Address & Default Contact Details
  End Enum

  Public Class Organisation

    Private mvStatusDateValid As Boolean
    Private mvSourceDateValid As Boolean
    Private mvOwnershipValid As Boolean
    Private mvAbbreviationValid As Boolean
    Private mvSortNameValid As Boolean
    Private mvDefaultContact As Contact
    Private mvCurrentAddress As Address
    Private mvDeletedContacts As List(Of Contact)
    Private mvSourceDesc As String = ""
    Private mvStatusDesc As String = ""
    Private mvStatusRgbValue As Integer
    Private mvWebAddressValid As Boolean
    Private mvWebAddress As String
    Private mvOwnershipGroupDesc As String
    Private mvPrincipalDepartmentDesc As String
    Private mvOwnershipAccessLevel As String
    Private mvOwnershipAccessLevelDesc As String
    Private mvOwners As CDBParameters
    Private mvStatusChangedAction As Integer
    Private mvAddressHistorical As Boolean

    Protected Overrides Sub ClearFields()
      'Add code here to clear any additional values held by the class
      mvStatusDateValid = False
      mvSourceDateValid = False
      mvOwnershipValid = False
      mvAbbreviationValid = False
      mvSortNameValid = False
      mvDefaultContact = Nothing
      mvCurrentAddress = Nothing
      mvDeletedContacts = Nothing
      mvSourceDesc = ""
      mvStatusDesc = ""
      mvWebAddress = ""
      mvWebAddressValid = False
      mvOwnershipGroupDesc = ""
      mvPrincipalDepartmentDesc = ""
      mvOwnershipAccessLevel = ""
      mvOwnershipAccessLevelDesc = ""
      mvOwners = Nothing 'BR15286
      mvStatusChangedAction = 0
      mvAddressHistorical = False
    End Sub

    Public Overloads Function GetRecordSetFields(ByVal pRSType As OrganisationRecordSetTypes) As String
      Dim vFields As String = ""
      Dim vAddressType As Address.AddressRecordSetTypes
      Dim vContactType As Contact.ContactRecordSetTypes

      If pRSType = OrganisationRecordSetTypes.ortAllFields Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "o")
      Else
        vFields = "o.organisation_number" 'ortNumber
        If (pRSType And OrganisationRecordSetTypes.ortName) > 0 Then vFields = vFields & ",o.contact_number,name,sort_name,abbreviation"
        If (pRSType And OrganisationRecordSetTypes.ortGroup) > 0 Then vFields = vFields & ",organisation_group"
        If (pRSType And OrganisationRecordSetTypes.ortPhone) > 0 Then vFields = vFields & ",o.dialling_code,o.std_code,o.telephone"
        If (pRSType And OrganisationRecordSetTypes.ortDetail) > 0 Then
          vFields = vFields & ",o.source,o.source_date,o.status,o.status_date,o.status_reason,o.department,o.notes,o.amended_on,o.amended_by"
          If mvClassFields(OrganisationFields.OwnershipGroup).InDatabase Then vFields = vFields & ",o.ownership_group"
          If mvClassFields(OrganisationFields.ResponseChannel).InDatabase Then vFields = vFields & ",o.response_channel"
        End If
        If (pRSType And OrganisationRecordSetTypes.ortDefaultContact) > 0 Then
          vContactType = Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName
          vFields = vFields & "," & DefaultContact.GetRecordSetFields(vContactType)
          If InStr(vFields, "o.contact_number,") > 0 Then vFields = Replace(vFields, "o.contact_number,", "")
        End If
        If (pRSType And OrganisationRecordSetTypes.ortAddress) > 0 Then
          vAddressType = Address.AddressRecordSetTypes.artNumber Or Address.AddressRecordSetTypes.artDetails
          If (pRSType And OrganisationRecordSetTypes.ortAddressCountry) > 0 Then vAddressType = vAddressType Or Address.AddressRecordSetTypes.artCountrySortCode
          vFields = vFields & "," & Address.GetRecordSetFields(vAddressType)
        End If
      End If
      Return vFields
    End Function

    Public Sub InitWithAddress(ByVal pEnv As CDBEnvironment, Optional ByRef pOrganisationNumber As Integer = 0, Optional ByRef pAddressNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vOrganisationType As OrganisationRecordSetTypes

      mvEnv = pEnv
      CheckClassFields()
      If pOrganisationNumber > 0 Then
        vOrganisationType = OrganisationRecordSetTypes.ortAll Or OrganisationRecordSetTypes.ortAddress Or OrganisationRecordSetTypes.ortAddressCountry
        vSQL = "SELECT " & GetRecordSetFields(vOrganisationType) & " FROM "
        If pAddressNumber > 0 Then
          vSQL = vSQL & "organisations o, organisation_addresses oa, addresses a, countries co WHERE o.organisation_number = " & pOrganisationNumber & " AND o.organisation_number = oa.organisation_number AND oa.address_number = " & pAddressNumber & " AND oa.address_number = a.address_number AND a.country = co.country"
        Else
          vSQL = vSQL & "organisations o, addresses a, countries co WHERE o.organisation_number = " & pOrganisationNumber & " AND a.address_number = o.address_number AND a.country = co.country"
        End If
        vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() Then
          InitFromRecordSet(pEnv, vRecordSet, vOrganisationType)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      ElseIf pAddressNumber > 0 Then
        'Address number only get organisation from address
        vOrganisationType = OrganisationRecordSetTypes.ortAll Or OrganisationRecordSetTypes.ortAddress Or OrganisationRecordSetTypes.ortAddressCountry
        vSQL = "SELECT " & GetRecordSetFields(vOrganisationType) & " FROM "
        vSQL = vSQL & "organisation_addresses oa, organisations o, addresses a, countries co WHERE oa.address_number = " & pAddressNumber & " AND oa.organisation_number = o.organisation_number AND oa.address_number = a.address_number AND a.country = co.country"
        vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() Then
          InitFromRecordSet(pEnv, vRecordSet, vOrganisationType)
        Else
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        SetDefaults()
      End If
    End Sub

    Public ReadOnly Property Address() As Address
      Get
        If mvCurrentAddress Is Nothing Then
          mvCurrentAddress = New Address(mvEnv)
          mvCurrentAddress.Init()
        End If
        Address = mvCurrentAddress
      End Get
    End Property

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As OrganisationRecordSetTypes)
      Dim vFields As CDBFields
      Dim vAddressType As Address.AddressRecordSetTypes
      Dim vContactType As Contact.ContactRecordSetTypes

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always grab the unique key, 'cos you need it for saving
        .SetItem(OrganisationFields.OrganisationNumber, vFields)
        If (pRSType And OrganisationRecordSetTypes.ortName) > 0 Then
          .SetItem(OrganisationFields.ContactNumber, vFields)
          .SetItem(OrganisationFields.Name, vFields)
          .SetItem(OrganisationFields.SortName, vFields)
          .SetItem(OrganisationFields.Abbreviation, vFields)
          mvAbbreviationValid = True
          mvSortNameValid = True
        End If
        If (pRSType And OrganisationRecordSetTypes.ortGroup) > 0 Then
          .SetItem(OrganisationFields.OrganisationGroup, vFields)
        End If
        If (pRSType And OrganisationRecordSetTypes.ortPhone) > 0 Then
          .SetItem(OrganisationFields.DiallingCode, vFields)
          .SetItem(OrganisationFields.StdCode, vFields)
          .SetItem(OrganisationFields.Telephone, vFields)
        End If
        If (pRSType And OrganisationRecordSetTypes.ortDetail) > 0 Then
          .SetItem(OrganisationFields.Source, vFields)
          .SetItem(OrganisationFields.SourceDate, vFields)
          .SetItem(OrganisationFields.Status, vFields)
          .SetItem(OrganisationFields.StatusDate, vFields)
          .SetItem(OrganisationFields.StatusReason, vFields)
          .SetItem(OrganisationFields.Department, vFields)
          .SetOptionalItem(OrganisationFields.OwnershipGroup, vFields)
          .SetOptionalItem(OrganisationFields.ResponseChannel, vFields)
          .SetItem(OrganisationFields.Notes, vFields)
          .SetItem(OrganisationFields.AmendedOn, vFields)
          .SetItem(OrganisationFields.AmendedBy, vFields)
          mvSourceDateValid = True
          mvStatusDateValid = True
        End If
        If pRSType = OrganisationRecordSetTypes.ortAllFields Then
          .SetItem(OrganisationFields.AddressNumber, vFields)
        Else
          If (pRSType And OrganisationRecordSetTypes.ortAddress) > 0 Then
            .SetItem(OrganisationFields.AddressNumber, vFields)
            vAddressType = Address.AddressRecordSetTypes.artNumber Or Address.AddressRecordSetTypes.artDetails
            If (pRSType And OrganisationRecordSetTypes.ortAddressCountry) > 0 Then vAddressType = vAddressType Or Address.AddressRecordSetTypes.artCountrySortCode
            Address.InitFromRecordSet(mvEnv, pRecordSet, vAddressType)
          End If
          If (pRSType And OrganisationRecordSetTypes.ortDefaultContact) > 0 Then
            .SetItem(OrganisationFields.ContactNumber, vFields)
            vContactType = Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName
            DefaultContact.InitFromRecordSet(mvEnv, pRecordSet, vContactType)
          End If
        End If
      End With
      mvOwnershipValid = False
    End Sub

    Public ReadOnly Property DefaultContact() As Contact
      Get
        If mvDefaultContact Is Nothing Then
          mvDefaultContact = New Contact(mvEnv)
          mvDefaultContact.Init()
        End If
        If mvDefaultContact.ContactNumber <> ContactNumber Then mvDefaultContact.Init(ContactNumber)
        DefaultContact = mvDefaultContact
      End Get
    End Property

    Public Sub ClearDefaultPhoneNumber()
      mvClassFields.Item(OrganisationFields.DiallingCode).Value = ""
      mvClassFields.Item(OrganisationFields.StdCode).Value = ""
      mvClassFields.Item(OrganisationFields.Telephone).Value = ""
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vMsg As String = DeleteOrganisation()
      If vMsg.Length > 0 Then Throw New Exception(vMsg)
    End Sub

    Public Function DeleteOrganisation() As String
      Dim vMsg As String

      mvDeletedContacts = New List(Of Contact)
      mvEnv.Connection.StartTransaction()
      vMsg = DeleteOrganisationData()
      If vMsg.Length = 0 Then
        mvEnv.Connection.CommitTransaction()
      Else
        mvEnv.Connection.RollbackTransaction()
      End If
      Return vMsg
    End Function

    Public Function CheckDeleteRights(Optional ByRef pCheckDefaultAddress As Boolean = True) As String
      Dim vRecordSet As CDBRecordSet
      Dim vConn As CDBConnection
      Dim vError As Boolean
      Dim vMsg As String = ""
      Dim vContact As New Contact(mvEnv)
      Dim vWhereFields As New CDBFields

      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipDepartments And Department <> mvEnv.User.Department Then
        Return (ProjectText.String16509) 'You must be in the Owner Department to delete this Organisation
      ElseIf mvEnv.User.HasItemAccessRights(CDBUser.AccessControlItems.aciContactDelete) = False Then
        Return (ProjectText.String17322) 'You do not have access to delete this Organisation
      End If
      vConn = mvEnv.Connection
      '--------------------------------------------------------------------------------------------------
      'Check for rights to delete from the communications log
      vRecordSet = vConn.GetRecordSet("SELECT public_delete, department_delete, department, creator_delete, created_by, communications_log_number FROM communications_log cl, document_classes dc WHERE address_number = " & AddressNumber & " AND dc.document_class = cl.document_class")
      With vRecordSet
        While .Fetch() = True And (vError = False)
          If ((.Fields(1).Bool) Or (.Fields(2).Bool And .Fields(3).Value = mvEnv.User.Department) Or (.Fields(4).Bool And .Fields(5).Value = mvEnv.User.Logname)) Then
            'OK to delete
          Else
            vMsg = String.Format(ProjectText.String16543, CStr(.Fields(6).IntegerValue)) 'No delete privilege for Communications Log number %s\r\n\r\nContact cannot be deleted
            vError = True
          End If
        End While
        .CloseRecordSet()
      End With
      '--------------------------------------------------------------------------------------------------
      'Check for rights to delete contacts
      If vError = False Then
        'Find all the contacts who have one of this organisations addresses as the default address
        vContact.Init()
        vRecordSet = vConn.GetRecordSet("SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtDefaultAddressNumber) & " FROM contacts c WHERE address_number IN (SELECT address_number FROM organisation_addresses WHERE organisation_number = " & OrganisationNumber & ")")
        With vRecordSet
          While .Fetch() = True And (vError = False)
            vContact.InitFromRecordSet(mvEnv, vRecordSet, Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtDefaultAddressNumber)
            'Check if this is the only address for this contact
            vWhereFields.Clear()
            vWhereFields.Add("contact_number", vContact.ContactNumber)
            vWhereFields.Add("address_number", vContact.AddressNumber, CDBField.FieldWhereOperators.fwoNotEqual)
            If vConn.GetCount("contact_addresses", vWhereFields) = 0 Then
              If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipDepartments And vContact.Department <> mvEnv.User.Department Then
                vMsg = (ProjectText.String16545) 'You must be in the owner department to delete Contacts\r\n\r\nOrganisation cannot be deleted
                vError = True
              ElseIf mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups And CDBEnvironment.GetOwnershipAccessLevel(vContact.OwnershipAccessLevel) <> CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite Then
                vMsg = (ProjectText.String17320) 'You must have Write access rights to delete Contacts & vbCrLf & vbCrLf & Organisation cannot be deleted
                vError = True
              Else
                vMsg = vContact.CheckDeleteRights(False)
                vError = vMsg.Length > 0
              End If
            Else
              If pCheckDefaultAddress = True Then vMsg = (ProjectText.String17321) 'Some Contacts have one of this Organisations Addresses as their Default Address and have other Addresses. You must change the Default Address of these Contacts before deleting this Organisation
              vError = vMsg.Length > 0
            End If
          End While
          .CloseRecordSet()
        End With
      End If
      Return vMsg
    End Function

    Private Function DeleteOrganisationData() As String
      Dim vRecordSet As CDBRecordSet
      Dim vConn As CDBConnection
      Dim vError As Boolean
      Dim vMsg As String = ""
      Dim vContact As New Contact(mvEnv)
      Dim vWhereFields As New CDBFields
      Dim vPostcode As String

      vConn = mvEnv.Connection
      vContact.Init()
      'Find all the contacts who have one of this organisations addresses as the default address
      vRecordSet = vConn.GetRecordSet("SELECT " & vContact.GetRecordSetFields(Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtDefaultAddressNumber) & " FROM contacts c WHERE address_number IN (SELECT address_number FROM organisation_addresses WHERE organisation_number = " & OrganisationNumber & ")")
      With vRecordSet
        While .Fetch() = True And (vError = False)
          vContact.InitFromRecordSet(mvEnv, vRecordSet, Contact.ContactRecordSetTypes.crtNumber Or Contact.ContactRecordSetTypes.crtName Or Contact.ContactRecordSetTypes.crtDetail Or Contact.ContactRecordSetTypes.crtDefaultAddressNumber)
          'Check if this is the only address for this contact
          vWhereFields.Clear()
          vWhereFields.Add("contact_number", vContact.ContactNumber)
          vWhereFields.Add("address_number", vContact.AddressNumber, CDBField.FieldWhereOperators.fwoNotEqual)
          If vConn.GetCount("contact_addresses", vWhereFields) = 0 Then
            vMsg = vContact.DeleteContact
            If vMsg.Length = 0 Then mvDeletedContacts.Add(vContact)
            vError = vMsg.Length > 0
          End If
        End While
        .CloseRecordSet()
      End With
      If vError Then Return vMsg
      '--------------------------------------------------------------------------------------------------
      'Delete from communications log history, communications log subjects and optionally communications_log_links
      '            where the log number matches the communications log
      vWhereFields.Clear()
      vWhereFields.Add("communications_log_number", CDBField.FieldTypes.cftLong, "SELECT cl.communications_log_number FROM communications_log cl WHERE cl.address_number = " & AddressNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.DeleteRecordsMultiTable("communications_log_history,communications_log_subjects,communications_log_links", vWhereFields)
      vWhereFields(1).Name = "unique_id"
      vWhereFields.Add("record_type", "D")
      vConn.DeleteRecords("sticky_notes", vWhereFields, False)
      '--------------------------------------------------------------------------------------------------
      'Delete from sticky_notes
      vWhereFields.Clear()
      vWhereFields.Add("unique_id", CDBField.FieldTypes.cftLong, OrganisationNumber)
      vWhereFields.Add("record_type", "O")
      vConn.DeleteRecords("sticky_notes", vWhereFields, False)
      '--------------------------------------------------------------------------------------------------
      'Delete from address_geographical_regions
      vWhereFields.Clear()
      vPostcode = vConn.GetValue("SELECT postcode FROM addresses WHERE address_number = " & AddressNumber)
      If Len(vPostcode) > 0 Then
        vWhereFields.Add("postcode", vPostcode, CDBField.FieldWhereOperators.fwoEqual)
        vWhereFields.Add("address_number",  AddressNumber, CDBField.FieldWhereOperators.fwoNotEqual)
        If vConn.GetCount("addresses", vWhereFields) = 0 Then
          vWhereFields.Remove((2))
          vConn.DeleteRecords("address_geographical_regions", vWhereFields, False)
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Delete from communications, communications log addresses contact_addresses contact_address_usages
      '            where the address_number matches
      vWhereFields.Clear()
      vWhereFields.Add("address_number", AddressNumber)
      vConn.DeleteRecordsMultiTable("communications,communications_log,addresses,contact_addresses,contact_address_usages", vWhereFields)
      '--------------------------------------------------------------------------------------------------
      'Delete from organisation_links
      vWhereFields.Clear()
      vWhereFields.Add("organisation_number_1", OrganisationNumber)
      vConn.DeleteRecords("organisation_links", vWhereFields, False)
      vWhereFields.Clear()
      vWhereFields.Add("organisation_number_2", OrganisationNumber)
      vConn.DeleteRecords("organisation_links", vWhereFields, False)
      '--------------------------------------------------------------------------------------------------
      'Delete from organisation_addresses, organisation_address_usages
      '            organisation_categories, contact_positions, contact_roles
      '            organisation_suppressions, organisation_users, organisations
      '            where the organisation_number matches
      vWhereFields.Clear()
      vWhereFields.Add("organisation_number", OrganisationNumber)
      vConn.DeleteRecordsMultiTable("organisation_addresses,organisation_address_usages,organisation_categories,contact_positions,contact_roles,organisation_suppressions,organisation_users,organisations", vWhereFields)

      '--------------------------------------------------------------------------------------------------
      'Delete from principal users
      vWhereFields.Clear()
      vWhereFields.Add("contact_number", OrganisationNumber)
      vConn.DeleteRecords("principal_users", vWhereFields, False)

      '--------------------------------------------------------------------------------------------------
      'Delete from department_notes
      vWhereFields.Clear()
      vWhereFields.Add("unique_id", OrganisationNumber)
      vWhereFields.Add("record_type", "O")
      vConn.DeleteRecords("department_notes", vWhereFields, False)
      Return vMsg
    End Function

    Public ReadOnly Property WebAddress() As String
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vSQL As String

        If Not mvWebAddressValid Then
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDevicesWWWAddress) Then
            vSQL = "SELECT " & mvEnv.Connection.DBSpecialCol("", "number") & " FROM communications co, devices d WHERE co.address_number = " & AddressNumber & " AND co.device = d.device AND www_address = 'Y'"
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then
              vSQL = vSQL & " AND is_active = 'Y' ORDER BY device_default DESC, co.device"
            Else
              vSQL = vSQL & " ORDER BY co.device"
            End If
            vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
            If vRecordSet.Fetch() Then
              mvWebAddress = vRecordSet.Fields(1).Value
            End If
            vRecordSet.CloseRecordSet()
          End If
          mvWebAddressValid = True
        End If
        Return mvWebAddress
      End Get
    End Property

    Public ReadOnly Property OwnershipAccessLevelDesc() As String
      Get
        GetOwnershipInfo()
        Return mvOwnershipAccessLevelDesc
      End Get
    End Property
    Public ReadOnly Property OwnershipAccessLevel() As String
      Get
        GetOwnershipInfo()
        Return mvOwnershipAccessLevel
      End Get
    End Property
    Public ReadOnly Property PrincipalDepartmentDesc() As String
      Get
        GetOwnershipInfo()
        Return mvPrincipalDepartmentDesc
      End Get
    End Property
    Public ReadOnly Property OwnershipGroupDesc() As String
      Get
        GetOwnershipInfo()
        Return mvOwnershipGroupDesc
      End Get
    End Property

    Private Sub GetOwnershipInfo()
      Dim vRecordSet As CDBRecordSet

      If Not mvOwnershipValid Then
        If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT ownership_group_desc,department_desc,oal.ownership_access_level,ownership_access_level_desc FROM ownership_groups og, departments d, ownership_group_users ogu, ownership_access_levels oal WHERE og.ownership_group = '" & OwnershipGroup & "' AND d.department = og.principal_department AND ogu.ownership_group = og.ownership_group AND ogu.logname = '" & mvEnv.User.Logname & "' AND ogu.valid_from " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (TodaysDate())) & " AND (ogu.valid_to " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (TodaysDate())) & " OR ogu.valid_to IS NULL) AND oal.ownership_access_level = ogu.ownership_access_level")
          If vRecordSet.Fetch() Then
            mvOwnershipGroupDesc = vRecordSet.Fields(1).Value
            mvPrincipalDepartmentDesc = vRecordSet.Fields(2).Value
            mvOwnershipAccessLevel = vRecordSet.Fields(3).Value
            mvOwnershipAccessLevelDesc = vRecordSet.Fields(4).Value
          End If
          vRecordSet.CloseRecordSet()
        Else
          If Owners.Exists(mvEnv.User.Department) Then
            mvOwnershipAccessLevel = CDBEnvironment.GetOwnershipAccessLevelCode(CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite)
          Else
            mvOwnershipAccessLevel = CDBEnvironment.GetOwnershipAccessLevelCode(CDBEnvironment.OwnershipAccessLevelTypes.oaltRead)
          End If
        End If
        mvOwnershipValid = True
      End If
    End Sub

    Public ReadOnly Property Owners() As CDBParameters
      Get
        If mvOwners Is Nothing Then InitOwners()
        Owners = mvOwners
      End Get
    End Property

    Private Sub InitOwners()
      mvOwners = New CDBParameters
      Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT cu.department, department_desc FROM organisation_users cu, departments d WHERE organisation_number = " & OrganisationNumber & " AND cu.department = d.department")
      With vRS
        While .Fetch()
          mvOwners.Add((.Fields(1).Value), .Fields(2).Value)
        End While
        .CloseRecordSet()
      End With
    End Sub

    Public Sub UpdateOwner()
      Dim vOriginalDepartment As String = mvClassFields.Item(OrganisationFields.Department).SetValue
      'We may have changed the owning department and need to update the contact_users records
      If Department <> vOriginalDepartment Then
        If Owners.Exists(Department) Then
          'We are just swapping ownership and will leave the original owner as a co-owner
        Else
          'We are adding a new department as the owner - remove the original (assuming it existed)
          AddOwner(Department)
          If Owners.Exists(vOriginalDepartment) Then RemoveOwner(Owners(vOriginalDepartment))
        End If
      End If
    End Sub

    Public Sub AddOwner(ByRef pOwner As String)
      Dim vFields As New CDBFields
      vFields.AddAmendedOnBy(mvEnv.User.Logname, TodaysDate)
      vFields.Add("organisation_number", OrganisationNumber)
      vFields.Add("department", pOwner)
      mvEnv.Connection.InsertRecord("organisation_users", vFields)
    End Sub

    Public Sub RemoveOwner(ByRef pOwner As CDBParameter)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("organisation_number", OrganisationNumber)
      vWhereFields.Add("department", pOwner.Name)
      mvEnv.Connection.DeleteRecords("organisation_users", vWhereFields)
    End Sub

    Public Sub SetDefaultAddress(ByRef pNewDefaultAddress As Address, ByRef pSetOldHistoric As Boolean)
      'This method is designed to set the default address of an organisation to the address
      'passed to this method
      'An assumption is made that the organisation class has been correctly initialised
      'and the current default address has been read
      Dim vOldAddressNumber As Integer = AddressNumber
      If pNewDefaultAddress.AddressNumber <> vOldAddressNumber Then

        'Now update the organisation
        mvClassFields.Item(OrganisationFields.AddressNumber).IntegerValue = pNewDefaultAddress.AddressNumber
        Save(mvEnv.User.Logname, True)
        If pSetOldHistoric Then
          Dim vContactAddress As New ContactAddress(mvEnv)
          With vContactAddress
            .InitFromContactAndAddress(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltOrganisation, OrganisationNumber, vOldAddressNumber)
            If .Existing Then
              .Historical = True
              If .ValidTo.Length = 0 Then .ValidTo = TodaysDate()
              .Save(mvEnv.User.Logname, True)
            End If
          End With
        End If
      End If
    End Sub

    Public ReadOnly Property StatusChangedAction() As Integer
      Get
        StatusChangedAction = mvStatusChangedAction
      End Get
    End Property

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      SetValid()
      Dim vCSN As New ContactSearchName
      Dim vNewRecord As Boolean
      If mvExisting Then
        If mvClassFields.Item(OrganisationFields.Name).ValueChanged Then vCSN.Update(mvEnv, OrganisationNumber, Name)
      Else
        vNewRecord = True
      End If
      SaveStatusHistory(pAmendedBy)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
      If vNewRecord Then vCSN.Create(mvEnv, OrganisationNumber, Name)
    End Sub

    Public Overloads Sub Save(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      Save(pAmendedBy, pAudit, 0)
    End Sub

    Private Sub SaveStatusHistory(Optional ByVal pAmendedBy As String = "")
      If mvClassFields.Item(OrganisationFields.Status).ValueChanged Then
        If mvExisting = True Then
          If mvEnv.GetConfigOption("cd_record_status_history") = True Then
            If pAmendedBy.Length = 0 Then pAmendedBy = mvEnv.User.Logname
            Dim vUpdateFields As New CDBFields
            vUpdateFields.AddAmendedOnBy(pAmendedBy)
            vUpdateFields.Add("contact_number", OrganisationNumber)
            vUpdateFields.Add("status", mvClassFields.Item(OrganisationFields.Status).SetValue)
            vUpdateFields.Add("status_reason", mvClassFields.Item(OrganisationFields.StatusReason).SetValue)
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbResponseChannel) Then vUpdateFields.Add("response_channel", mvClassFields.Item(OrganisationFields.ResponseChannel).SetValue)
            Dim vEndDate As String = mvClassFields.Item(OrganisationFields.StatusDate).Value
            If vEndDate.Length = 0 Then vEndDate = TodaysDate()
            vUpdateFields.Add("valid_to", CDBField.FieldTypes.cftDate, vEndDate)
            mvEnv.Connection.InsertRecord("status_history", vUpdateFields)
          End If
          Dim vActionSet As New ActionSet
          mvStatusChangedAction = vActionSet.CreateStatusChangeAction(mvEnv, OrganisationGroupCode, OrganisationNumber, (mvClassFields(OrganisationFields.Status).SetValue), Status)
        End If
      End If
    End Sub

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters, ByRef pNoCapitalisation As Boolean)
      Dim vContact As New Contact(mvEnv)
      Dim vCA As New ContactAddress(mvEnv)
      Dim vCP As New ContactPosition(mvEnv)
      Dim vPU As New PrincipalUser
      Dim vAbbreviation As String
      Dim vSortName As String
      Dim vGroup As String

      Init()

      If pParams.HasValue("OrganisationNumber") Then
        'Contact number is being passed.  E.g. via OpenLink
        ValidateOrganisationNumber(MaintenanceTypes.Insert, pParams("OrganisationNumber").Value)
        SetOrganisationNumber(pParams("OrganisationNumber").IntegerValue)
      End If

      vContact.Init()
      vContact.ContactType = Contact.ContactTypes.ctcOrganisation 'This is just setup for the Address Create
      Address.Create(pEnv, vContact, pParams, pParams("Address").Value = " ") ' pHouseName, pAddress, pTown, pCounty, pPostCode, pCountry, pPaf, pBranch, pBuildingNumber

      Name = pParams("Name").CapitalisedValue(pNoCapitalisation)
      vAbbreviation = pParams.ParameterExists("Abbreviation").Value
      If vAbbreviation.Length > 0 Then Abbreviation = vAbbreviation
      vSortName = pParams.ParameterExists("SortName").Value
      If vSortName.Length > 0 Then SortName = vSortName
      vGroup = pParams.ParameterExists("OrganisationGroup").Value
      If vGroup.Length = 0 Then vGroup = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtOrganisation).EntityGroupCode
      mvClassFields.Item(OrganisationFields.OrganisationGroup).Value = vGroup
      Notes = pParams.ParameterExists("Notes").Value
      Source = pParams("Source").Value
      SourceDate = pParams.OptionalValue("SourceDate", (TodaysDate()))

      If pParams.HasValue("Salutation") Then vContact.Salutation = pParams("Salutation").Value
      If pParams.HasValue("LabelName") Then vContact.LabelName = pParams("LabelName").Value
      If pParams.HasValue("Department") Then Department = pParams("Department").Value

      Status = pParams.ParameterExists("Status").Value
      If pParams.HasValue("StatusDate") Then StatusDate = pParams("StatusDate").Value
      If pParams.Exists("StatusReason") Then StatusReason = pParams("StatusReason").Value
      If pParams.HasValue("ResponseChannel") Then ResponseChannel = pParams("ResponseChannel").Value

      If pParams.HasValue("OwnershipGroup") Then
        OwnershipGroup = pParams("OwnershipGroup").Value
      Else
        OwnershipGroup = mvEnv.User.OwnershipGroup
      End If
      vContact.VATCategory = pParams.OptionalValue("VatCategory", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefOrgVatCat))

      'Try and get all the control numbers outside of the transaction
      Address.SetControlNumber()
      MyBase.SetControlNumber()
      mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnContact, 1)
      mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnAddressLink, 2)
      mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnPosition, 1)

      mvEnv.Connection.StartTransaction()
      Address.Save(mvEnv.User.UserID, True)
      mvClassFields.Item(OrganisationFields.AddressNumber).IntegerValue = Address.AddressNumber
      Save(mvEnv.User.UserID, True)
      AddUser(Department, True)
      Dim vAddressValidFrom As String = TodaysDate()
      If pParams.ContainsKey("ValidFrom") Then vAddressValidFrom = pParams("ValidFrom").Value
      Dim vAddressValidTo As String = ""
      If pParams.ContainsKey("ValidTo") Then vAddressValidTo = pParams("ValidTo").Value
      vCA.Create(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltOrganisation, OrganisationNumber, AddressNumber, "N", vAddressValidFrom, vAddressValidTo)
      vCA.Save(mvEnv.User.UserID, True)
      AddDummyContact(vContact)
      DefaultContact.AddUser(Department, True)
      vCA = New ContactAddress(mvEnv)
      vCA.Create(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, OrganisationNumber, AddressNumber, "N", vAddressValidFrom, vAddressValidTo)
      vCA.Save(mvEnv.User.UserID, True)
      vCP.Create(OrganisationNumber, AddressNumber, OrganisationNumber, "Y", "Y")
      vCP.Save(mvEnv.User.UserID, True)
      If pParams.HasValue("PrincipalUser") Then vPU.Create(mvEnv, OrganisationNumber, pParams("PrincipalUser").Value, pParams.ParameterExists("PrincipalUserReason").Value)
      If pParams.HasValue("VatNumber") Then vContact.VATNumber = pParams("VatNumber").Value
      mvEnv.Connection.CommitTransaction()
    End Sub

    Private Sub ValidateOrganisationNumber(pType As MaintenanceTypes, vNumber As String)
      Dim vOrganisationNumber As Integer
      If Integer.TryParse(vNumber, vOrganisationNumber) Then
        'First check for uniqueness if inserting
        If pType = MaintenanceTypes.Insert Then
          Dim vDup As New Contact(Me.Environment) 'Must not be Org record as it could be number already used by Contact
          vDup.Init(vOrganisationNumber)
          If vDup.Existing Then
            RaiseError(DataAccessErrors.daeRecordExists, ProjectText.LangOrganisationNumber)
          End If
        End If
        'Now check for boundaries
        Dim vSQL As New SQLStatement(Me.Environment.Connection, "control_number", "control_numbers", New CDBField("control_number_type", Me.ClassFields.ControlNumberType))
        Dim vMaxNumber As Integer = CType(Me.Environment.Connection.GetValue(vSQL.SQL, True), Integer)
        vMaxNumber -= 1 'the last created org number should be the control number minus one.  The next Org number is the Control Number
        If vOrganisationNumber > vMaxNumber Then
          RaiseError(DataAccessErrors.daeRangeExceeded, ProjectText.LangOrganisationNumber, vMaxNumber.ToString())
        End If
      Else
        'The code here will only be reached when a Contact is created with parameters that have bypassed ValidateParameterList, which should be never.
        RaiseError(DataAccessErrors.daeInvalidParameter, ProjectText.LangOrganisationNumber)
      End If
    End Sub

    Public Sub SetOrganisationNumber(ByVal pOrganisationNumber As Integer)
      mvClassFields(OrganisationFields.OrganisationNumber).IntegerValue = pOrganisationNumber
    End Sub

    Public Sub CreateAddress(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      Dim vContact As New Contact(mvEnv)
      Dim vAddress As New Address(mvEnv)
      Dim vCA As New ContactAddress(mvEnv)

      vContact.Init()
      vContact.ContactType = Contact.ContactTypes.ctcOrganisation 'This is just setup for the Address Create
      vAddress.Init()
      vAddress.Create(pEnv, vContact, pParams) '.ParameterExists("HouseName").Value, pParams("Address").Value, pParams("Town").Value, pParams.ParameterExists("County").Value, pParams.ParameterExists("Postcode").Value, pParams("Country").Value, pParams.ParameterExists("PafStatus").Value, pParams.ParameterExists("Branch").Value, pParams.ParameterExists("BuildingNumber").Value
      vAddress.Save(pEnv.User.UserID, True)
      Dim vAddressValidFrom As String = TodaysDate()
      If pParams.ContainsKey("ValidFrom") Then vAddressValidFrom = pParams("ValidFrom").Value
      Dim vAddressValidTo As String = ""
      If pParams.ContainsKey("ValidTo") Then vAddressValidTo = pParams("ValidTo").Value
      vCA.Create(pEnv, ContactAddress.ContactAddresssLinkTypes.caltOrganisation, OrganisationNumber, vAddress.AddressNumber, "N", vAddressValidFrom, vAddressValidTo)
      vCA.Save(mvEnv.User.UserID, True)
      mvCurrentAddress = vAddress
    End Sub

    Public Sub AddUser(ByVal pDepartment As String, Optional ByVal pNew As Boolean = False)
      Dim vFields As New CDBFields
      Dim vAdd As Boolean

      vFields.Add("organisation_number", OrganisationNumber)
      vFields.Add("department", pDepartment)
      If pNew Then
        vAdd = True
      Else
        vAdd = mvEnv.Connection.GetCount("organisation_users", vFields) = 0
      End If
      If vAdd Then
        vFields.AddAmendedOnBy(mvEnv.User.UserID)
        mvEnv.Connection.InsertRecord("organisation_users", vFields)
      End If
    End Sub

    Public Sub AddDummyContact(Optional ByRef pContact As Contact = Nothing)
      Dim vLabelName As String

      If pContact Is Nothing Then
        pContact = New Contact(mvEnv)
        pContact.Init()
      End If

      With pContact
        If .LabelName.Length = 0 Then
          vLabelName = TruncateString(Name, Contact.MaxLabelNameLength).Trim
        Else
          vLabelName = .LabelName
        End If
        .ContactType = Contact.ContactTypes.ctcOrganisation
        If .Salutation.Length = 0 Then
          .Salutation = mvEnv.GetConfig("default_org_salutation")
          If .Salutation.Length = 0 Then .Salutation = "Dear Sirs"
        End If
        If .VATCategory.Length = 0 Then .VATCategory = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefOrgVatCat)
        .ContactGroupCode = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode
        .SetContactNumber(OrganisationNumber)
        .AddressNumber = AddressNumber
        .Surname = TruncateString(Name, 50).Trim  'BR18969
        .Sex = Contact.ContactSex.cscUnknown
        .Source = Source
        .Status = Status
        .StatusDate = StatusDate
        .StatusReason = StatusReason
        .SourceDate = SourceDate
        .Department = Department
        .OwnershipGroup = OwnershipGroup
        .LabelName = vLabelName
        .Save()
        mvDefaultContact = pContact
      End With
    End Sub

    Public Function DedupOrganisation(ByVal pName As String, Optional ByRef pContactNo As Integer = 0) As Boolean
      'Attempt to find an organisation from the given name (the address is unknown!!)
      'If the contact number is passed, first check the contact_positions
      Dim vColl As New CollectionList(Of Organisation)
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vFound As Boolean
      Dim vOrg As Organisation

      InitClassFields()
      pName = StrConv(pName, VbStrConv.ProperCase)
      'First see if our contact has a position at our organisation
      If pContactNo > 0 Then
        vSQL = "SELECT DISTINCT " & GetRecordSetFields(OrganisationRecordSetTypes.ortName) & " FROM contact_positions cp, organisations o"
        vSQL = vSQL & " WHERE cp.contact_number = " & pContactNo & " AND o.organisation_number = cp.organisation_number"
        vSQL = vSQL & " ORDER BY o.organisation_number"
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        While vRS.Fetch
          vOrg = New Organisation(mvEnv)
          vOrg.InitFromRecordSet(mvEnv, vRS, OrganisationRecordSetTypes.ortName)
          vColl.Add(vRS.Fields("organisation_number").Value, vOrg)
        End While
        vRS.CloseRecordSet()

        For Each vOrg In vColl
          If StrComp(pName, vOrg.Name, CompareMethod.Text) = 0 Then
            With vOrg
              'Set up this class with the name information
              mvClassFields(OrganisationFields.OrganisationNumber).IntegerValue = .OrganisationNumber
              ContactNumber = .ContactNumber
              Name = .Name
              SortName = .SortName
              Abbreviation = .Abbreviation
            End With
            mvExisting = True
            Return True
            Exit For
          End If
        Next vOrg
        vColl = New CollectionList(Of Organisation)
      End If

      'If nothing found then just dedup the organisation name
      If Not vFound Then
        vSQL = "SELECT " & GetRecordSetFields(OrganisationRecordSetTypes.ortName) & " FROM organisations o"
        vSQL = vSQL & " WHERE name " & mvEnv.Connection.DBLike("*" & pName & "*", CDBField.FieldTypes.cftUnicode)
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        While vRS.Fetch
          vOrg = New Organisation(mvEnv)
          vOrg.InitFromRecordSet(mvEnv, vRS, OrganisationRecordSetTypes.ortName)
          vColl.Add(vRS.Fields("organisation_number").Value, vOrg)
        End While
        vRS.CloseRecordSet()

        If vColl.Count() > 1 Then
          'More than 1 matching organisation, so can not match
          vFound = True
        ElseIf vColl.Count() = 1 Then
          'Only the 1 organisation
          vOrg = vColl.Item(0)
          With vOrg
            'Set up this class with the name information
            mvClassFields(OrganisationFields.OrganisationNumber).IntegerValue = .OrganisationNumber
            ContactNumber = .ContactNumber
            Name = .Name
            SortName = .SortName
            Abbreviation = .Abbreviation
          End With
          mvExisting = True
          Return True
        End If
      End If
      Return vFound
    End Function

  End Class
End Namespace
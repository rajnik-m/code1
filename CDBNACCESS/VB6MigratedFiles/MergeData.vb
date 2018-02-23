

Namespace Access
  Public Class MergeData

    Public Sub ContactMerge(ByRef pEnv As CDBEnvironment, ByRef pJob As JobSchedule, ByRef pParams As CDBParameters)
      Dim vPAddress As Integer
      Dim vDAddress As Integer
      Dim vPContact As New Contact(pEnv)
      Dim vDContact As New Contact(pEnv)
      Dim vPMemStatus As Contact.ContactMemberStatuses
      Dim vDMemStatus As Contact.ContactMemberStatuses
      Dim vFields As New CDBFields

      Dim vPrimaryContactNumber As Integer
      Dim vPrimaryAddressNumber As Integer
      Dim vDuplicateContactNumber As Integer
      Dim vDuplicateAddressNumber As Integer
      Dim vQueue As Boolean
      Try

        vPrimaryContactNumber = pParams("ContactNumber").IntegerValue
        vPrimaryAddressNumber = pParams("AddressNumber").IntegerValue
        vDuplicateContactNumber = pParams("DuplicateContactNumber").IntegerValue
        vDuplicateAddressNumber = pParams("DuplicateAddressNumber").IntegerValue
        vQueue = pParams("Queue").Bool

        If vPrimaryContactNumber = vDuplicateContactNumber Then
          RaiseError(DataAccessErrors.daeMergeSameContact) 'You may not merge a contact with itself
        Else
          vPContact.Init(vPrimaryContactNumber)
          vDContact.Init(vDuplicateContactNumber)
          If vPContact.ContactNumber > 0 And vDContact.ContactNumber > 0 Then
            CheckOwnershipValid(pEnv, vPContact, vDContact)
            vPMemStatus = vPContact.ContactMemberStatus
            vDMemStatus = vDContact.ContactMemberStatus
            If (vPMemStatus = Contact.ContactMemberStatuses.cmsSingle) And (vDMemStatus = Contact.ContactMemberStatuses.cmsSingle) Then
              RaiseError(DataAccessErrors.daeMergeMembershipTypes) 'These contacts have current, mutually exclusive membership types and can not be merged
            End If
            If CheckExternalContact(pEnv, vDuplicateContactNumber) Then
              RaiseError(DataAccessErrors.daeMergeExternal) 'The duplicate contact exists in an external database and cannot be merged
            End If
            If CheckLegacies(pEnv, vPContact.ContactNumber, vDContact.ContactNumber) Then
              RaiseError(DataAccessErrors.daeMergeLegacy) 'These contacts each have a legacy and may not be merged
            End If
            If (vPContact.ContactType = Contact.ContactTypes.ctcContact And vDContact.ContactType = Contact.ContactTypes.ctcJoint) Or (vPContact.ContactType = Contact.ContactTypes.ctcJoint And vDContact.ContactType = Contact.ContactTypes.ctcContact) Then
              RaiseError(DataAccessErrors.daeMergeJointIndividual) 'A Joint Contact cannot be merged with an Individual Contact
            End If
            If vPContact.ContactType = Contact.ContactTypes.ctcJoint And vDContact.ContactType = Contact.ContactTypes.ctcJoint Then
              If pParams.OptionalValue("ConfirmJointMerge", "N") = "N" Then RaiseError(DataAccessErrors.daeMergeJointToJoint)
            End If
            If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIrishGiftAid) = True Then
              If CheckAppropriateCertificates(pEnv, vPContact.ContactNumber, vDContact.ContactNumber) Then
                RaiseError(DataAccessErrors.daeMergeCertificates) 'These contacts have live certificates and cannot be merged
              End If
            End If
            If pParams.OptionalValue("ConfirmMergeCreditCustomers", "N") = "N" Then
              If CreditCustomerCheck(pEnv, vPrimaryContactNumber, vDuplicateContactNumber) Then
                RaiseError(DataAccessErrors.daeMergeCreditCustomers)
              End If
            End If

            If vQueue Then
              If pEnv.Connection.GetCount("duplicate_contacts", Nothing, "contact_number_1 IN(" & vPContact.ContactNumber & "," & vDContact.ContactNumber & ")") > 0 Or pEnv.Connection.GetCount("duplicate_contacts", Nothing, "contact_number_2 IN(" & vPContact.ContactNumber & "," & vDContact.ContactNumber & ")") > 0 Then
                RaiseError(DataAccessErrors.daeMergeAlreadySet) 'One of these Contacts has already been set for DeDuplication
              Else
                If pParams.OptionalValue("Confirm", "N") = "N" Then RaiseError(DataAccessErrors.daeMergeQueueConfirm)
                With vFields
                  'This is the opposite way round to what you might expect!!
                  .Add("contact_number_1", CDBField.FieldTypes.cftLong, vDContact.ContactNumber)
                  .Add("address_number_1", CDBField.FieldTypes.cftLong, vDuplicateAddressNumber)
                  .Add("contact_number_2", CDBField.FieldTypes.cftLong, vPContact.ContactNumber)
                  .Add("address_number_2", CDBField.FieldTypes.cftLong, vPrimaryAddressNumber)
                  .Add("match_or_potential", CDBField.FieldTypes.cftCharacter, "M")
                  .Add("application_code", CDBField.FieldTypes.cftCharacter, "CM")
                  .Add("run_date", CDBField.FieldTypes.cftDate, TodaysDate)
                End With
                pEnv.Connection.InsertRecord("duplicate_contacts", vFields)
              End If
            Else
              If vPrimaryAddressNumber <> vDuplicateAddressNumber Then
                If Not pParams.Exists("DeleteAddress") Then RaiseError(DataAccessErrors.daeMergeDeleteAddress)
                If pParams("DeleteAddress").Bool Then
                  'BR21369 : Before Deleteing Default Address on the Duplicate, if it is an Org Address on any of the Contact Positions.
                  'In which case it should not be deleted to prevent a broken link
                  Dim vDeleteDupDefAddr As Boolean
                  Dim vWhereFields As New CDBFields
                  If vDuplicateAddressNumber > 0 Then
                    vWhereFields.Add("contact_number", CDBField.FieldTypes.cftCharacter, vDContact.ContactNumber)
                    vWhereFields.Add("address_number", CDBField.FieldTypes.cftCharacter, vDuplicateAddressNumber)
                    vDeleteDupDefAddr = Not CBool(pEnv.Connection.GetCount("contact_positions", vWhereFields))
                    vWhereFields.Clear()
                    If Not vDeleteDupDefAddr Then
                      RaiseError(DataAccessErrors.daeMergeDeleteAddress)
                    Else
                      vPAddress = vPrimaryAddressNumber
                      vDAddress = vDuplicateAddressNumber
                    End If
                  End If
                End If
              End If
              If pParams.OptionalValue("Confirm", "N") = "N" Then RaiseError(DataAccessErrors.daeMergeConfirm)
              vPContact.DoContactMerge(pJob, pEnv.Connection, vDContact, vPAddress, vDAddress, pParams("Notes").Value)
            End If
          Else
            'Cannot find the contacts?
          End If
        End If
      Catch ex As Advanced.Data.Merge.MergeObjectException
        RaiseError(DataAccessErrors.daeMergeObjectException, ProjectText.LangContact, ex.Message)
      End Try

    End Sub

    Public Sub OrganisationMerge(ByRef pEnv As CDBEnvironment, ByRef pJob As JobSchedule, ByRef pParams As CDBParameters)
      Dim vPContact As New Contact(pEnv)
      Dim vDContact As New Contact(pEnv)
      Dim vDeleteDupAddress As Boolean
      Dim vFields As New CDBFields
      Dim vPMemStatus As Contact.ContactMemberStatuses
      Dim vDMemStatus As Contact.ContactMemberStatuses

      Dim vPrimaryContactNumber As Integer
      Dim vPrimaryAddressNumber As Integer
      Dim vDuplicateContactNumber As Integer
      Dim vDuplicateAddressNumber As Integer
      Dim vQueue As Boolean

      vPrimaryContactNumber = pParams("ContactNumber").IntegerValue
      vPrimaryAddressNumber = pParams("AddressNumber").IntegerValue
      vDuplicateContactNumber = pParams("DuplicateContactNumber").IntegerValue
      vDuplicateAddressNumber = pParams("DuplicateAddressNumber").IntegerValue
      vQueue = pParams("Queue").Bool

      If vPrimaryContactNumber = vDuplicateContactNumber Then
        RaiseError(DataAccessErrors.daeMergeSameOrganisation) 'You may not merge an organisation with itself
      Else
        If pEnv.Connection.GetCount("branches", Nothing, "organisation_number IN (" & vPrimaryContactNumber & "," & vDuplicateContactNumber & ")") > 1 Then
          RaiseError(DataAccessErrors.daeMergeBothBranches) 'Both organisations are branches
        Else
          vPContact.Init(vPrimaryContactNumber)
          vDContact.Init(vDuplicateContactNumber)
          If vPContact.ContactNumber > 0 And vDContact.ContactNumber > 0 Then
            CheckOwnershipValid(pEnv, vPContact, vDContact)
            vPMemStatus = vPContact.ContactMemberStatus
            vDMemStatus = vDContact.ContactMemberStatus
            If (vPMemStatus = Contact.ContactMemberStatuses.cmsSingle) And (vDMemStatus = Contact.ContactMemberStatuses.cmsSingle) Then
              RaiseError(DataAccessErrors.daeMergeOrgMembershipTypes) 'These organisations have current, mutually exclusive membership types and can not be merged
            Else
              If pParams.OptionalValue("ConfirmMergeCreditCustomers", "N") = "N" Then
                If CreditCustomerCheck(pEnv, vPrimaryContactNumber, vDuplicateContactNumber) Then
                  RaiseError(DataAccessErrors.daeMergeCreditCustomers)
                End If
              End If
              If vQueue Then
                If pEnv.Connection.GetCount("duplicate_contacts", Nothing, "contact_number_1 IN(" & vPContact.ContactNumber & "," & vDContact.ContactNumber & ")") > 0 Or pEnv.Connection.GetCount("duplicate_contacts", Nothing, "contact_number_2 IN(" & vPContact.ContactNumber & "," & vDContact.ContactNumber & ")") > 0 Then
                  RaiseError(DataAccessErrors.daeMergeOrgAlreadyset) 'One of these Organisations has already been set for DeDuplication
                Else
                  If pParams.OptionalValue("Confirm", "N") = "N" Then RaiseError(DataAccessErrors.daeMergeQueueConfirm)
                  With vFields
                    'This is the opposite way round to what you might expect!!
                    .Add("contact_number_1", CDBField.FieldTypes.cftLong, vDContact.ContactNumber)
                    .Add("address_number_1", CDBField.FieldTypes.cftLong, vDuplicateAddressNumber)
                    .Add("contact_number_2", CDBField.FieldTypes.cftLong, vPContact.ContactNumber)
                    .Add("address_number_2", CDBField.FieldTypes.cftLong, vPrimaryAddressNumber)
                    .Add("match_or_potential", CDBField.FieldTypes.cftCharacter, "M")
                    .Add("application_code", CDBField.FieldTypes.cftCharacter, "OM")
                    .Add("run_date", CDBField.FieldTypes.cftDate, TodaysDate)
                  End With
                  pEnv.Connection.InsertRecord("duplicate_contacts", vFields)
                End If
              Else
                If Not pParams.Exists("DeleteAddress") Then RaiseError(DataAccessErrors.daeMergeDeleteAddress)
                vDeleteDupAddress = pParams("DeleteAddress").Bool
                If pParams.OptionalValue("Confirm", "N") = "N" Then RaiseError(DataAccessErrors.daeMergeConfirm)
                vPContact.DoOrgMerge(pJob, pEnv.Connection, vDContact, vPrimaryAddressNumber, vDuplicateAddressNumber, pParams("Notes").Value, vDeleteDupAddress)
              End If
            End If
          End If
        End If
      End If
    End Sub

    Private Sub CheckOwnershipValid(ByRef pEnv As CDBEnvironment, ByRef pPrimaryContact As Contact, ByRef pDuplicateContact As Contact)
      Dim vMsg As String = ""
      Dim vPrimaryOrg As Organisation
      Dim vDuplicateOrg As Organisation
      If pEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        If pPrimaryContact.ContactType = Contact.ContactTypes.ctcOrganisation And pDuplicateContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
          vPrimaryOrg = New Organisation(pEnv)
          vPrimaryOrg.Init((pPrimaryContact.ContactNumber))
          vDuplicateOrg = New Organisation(pEnv)
          vDuplicateOrg.Init((pDuplicateContact.ContactNumber))
          If pEnv.User.OwnershipAccessLevel(vPrimaryOrg.OrganisationNumber, pEnv.EntityGroups.GroupFromCode(EntityGroup.EntityGroupTypes.egtOrganisation, (vPrimaryOrg.OrganisationGroupCode)), vMsg) <> CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite Or pEnv.User.OwnershipAccessLevel(vDuplicateOrg.OrganisationNumber, pEnv.EntityGroups.GroupFromCode(EntityGroup.EntityGroupTypes.egtContact, (vDuplicateOrg.OrganisationGroupCode)), vMsg) <> CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite Then
            RaiseError(DataAccessErrors.daeMergeOwnership)
          End If
        Else
          If pEnv.User.OwnershipAccessLevel(pPrimaryContact.ContactNumber, pEnv.EntityGroups.GroupFromCode(EntityGroup.EntityGroupTypes.egtContact, (pPrimaryContact.ContactGroupCode)), vMsg) <> CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite Or pEnv.User.OwnershipAccessLevel(pDuplicateContact.ContactNumber, pEnv.EntityGroups.GroupFromCode(EntityGroup.EntityGroupTypes.egtContact, (pDuplicateContact.ContactGroupCode)), vMsg) <> CDBEnvironment.OwnershipAccessLevelTypes.oaltWrite Then
            RaiseError(DataAccessErrors.daeMergeOwnership)
          End If
        End If
      End If
    End Sub

    Private Function CheckExternalContact(ByRef pEnv As CDBEnvironment, ByVal pContactNumber As Integer) As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vRecordSet2 As CDBRecordSet
      Dim vSQL As String

      'Check referential integrity on custom forms
      If pEnv.GetConfigOption("option_custom_data", False) Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT referential_sql, db_name FROM custom_forms WHERE client = '" & pEnv.ClientCode & "' AND custom_form >= " & pEnv.FirstCustomFormNumber & " AND custom_form <= " & pEnv.LastCustomFormNumber & " AND referential_sql IS NOT NULL")
        While vRecordSet.Fetch() = True
          vSQL = vRecordSet.Fields(1).Value
          vSQL = ReplaceString(vSQL, "?", CStr(pContactNumber))
          vSQL = ReplaceString(vSQL, "#", CStr(pContactNumber))
          vRecordSet2 = pEnv.GetConnection((vRecordSet.Fields(2).Value)).GetRecordSet(vSQL)
          If vRecordSet2.Fetch() = True Then CheckExternalContact = True
          vRecordSet2.CloseRecordSet()
        End While
        vRecordSet.CloseRecordSet()
      End If
    End Function
    ''' <summary>
    ''' This method finds out, if there are any duplicate exam_units between main contact and duplicate contact. If yes then we cannot merge the record
    ''' </summary>
    ''' <param name="pEnv">Current environment class</param>
    ''' <param name="pMainContact">Main conta</param>
    ''' <param name="pDuplicateContact"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckDuplicateExamUnitRecords(ByRef pEnv As CDBEnvironment, ByVal pMainContact As Integer, ByVal pDuplicateContact As Integer) As Boolean
      Dim SubWhereClause As New CDBFields(New CDBField("contact_number", pMainContact))
      Dim vSubSql As New SQLStatement(pEnv.Connection, "exam_unit_id", "exam_student_header", SubWhereClause)
      Dim vMainWhereClause As New CDBFields(New CDBField("contact_number", pDuplicateContact))
      vMainWhereClause.Add("exam_unit_id", String.Format("{0}", vSubSql.SQL), CDBField.FieldWhereOperators.fwoIn)
      Dim vSqlStatement As New SQLStatement(pEnv.Connection, "exam_unit_id", "exam_student_header", vMainWhereClause)
      Dim vDT As DataTable = pEnv.Connection.GetDataTable(vSqlStatement)

      If vDT.Rows.Count > 0 Then Return True
      Return False
    End Function

    Private Function CheckLegacies(ByRef pEnv As CDBEnvironment, ByVal pPContactNumber As Integer, ByVal pDContactNumber As Integer) As Boolean
      Dim vPRecordSet As CDBRecordSet
      Dim vDRecordSet As CDBRecordSet
      Dim vPCountLegacy As Integer
      Dim vDCountLegacy As Integer

      vPRecordSet = pEnv.Connection.GetRecordSet("SELECT count(*)  AS  p_record_count FROM contact_legacies WHERE contact_number = " & pPContactNumber)
      If vPRecordSet.Fetch() = True Then
        vPCountLegacy = CInt(vPRecordSet.Fields("p_record_count").Value)
      End If
      vPRecordSet.CloseRecordSet()
      vDRecordSet = pEnv.Connection.GetRecordSet("SELECT count(*)  AS  d_record_count FROM contact_legacies WHERE contact_number = " & pDContactNumber)
      If vDRecordSet.Fetch() = True Then
        vDCountLegacy = CInt(vDRecordSet.Fields("d_record_count").Value)
      End If
      vDRecordSet.CloseRecordSet()
      CheckLegacies = vPCountLegacy > 0 And vDCountLegacy > 0
    End Function

    Private Function CheckAppropriateCertificates(ByRef pEnv As CDBEnvironment, ByVal pOriginalContact As Integer, ByVal pDuplicateContact As Integer) As Boolean
      'Check if both contacts have Live Appropriate Certificates
      'Return True if Live Certificates are found
      Dim vOrigRS As CDBRecordSet
      Dim vDupRS As CDBRecordSet
      Dim vOrigAP As GaAppropriateCertificate
      Dim vDupAP As GaAppropriateCertificate
      Dim vNewAPs As Collection
      Dim vOldAPs As Collection
      Dim vSQL As String

      'Get all certificates from Primary Contact
      vSQL = "SELECT * FROM ga_appropriate_certificates gac WHERE contact_number = " & pOriginalContact & " ORDER BY start_date"
      vOrigRS = pEnv.Connection.GetRecordSet(vSQL)

      vOrigAP = New GaAppropriateCertificate(pEnv)
      vNewAPs = New Collection

      While vOrigRS.Fetch() = True
        vOrigAP.InitFromRecordSet(vOrigRS)
        vNewAPs.Add(vOrigAP)
      End While
      vOrigRS.CloseRecordSet()

      If vNewAPs.Count() > 0 Then

        'Get all certificates of Second contact having same start_date
        vSQL = "SELECT * FROM ga_appropriate_certificates gac WHERE contact_number = " & pDuplicateContact & " ORDER BY start_date"
        vDupRS = pEnv.Connection.GetRecordSet(vSQL)

        vDupAP = New GaAppropriateCertificate(pEnv)
        vOldAPs = New Collection

        While vDupRS.Fetch() = True
          vDupAP.InitFromRecordSet(vDupRS)
          vOldAPs.Add(vDupAP)
        End While
        vDupRS.CloseRecordSet()

        If vOldAPs.Count() > 0 Then
          For Each vOrigAP In vNewAPs 'Check against each certificate and return TRUE if both contacts have a live certificate for a same year
            For Each vDupAP In vOldAPs
              If CDate(vOrigAP.StartDate) = CDate(vDupAP.StartDate) Then
                If Len(vOrigAP.CancellationReason) = 0 And Len(vDupAP.CancellationReason) = 0 Then
                  CheckAppropriateCertificates = True
                  Exit Function
                End If
              End If
            Next vDupAP
          Next vOrigAP
        End If
      End If
      CheckAppropriateCertificates = False
    End Function

    Function CreditCustomerCheck(ByRef pEnv As CDBEnvironment, ByVal pOriginalContact As Integer, ByVal pDuplicateContact As Integer) As Boolean
      'special case for credit customers: index is unique on contact_number, company, sales_ledger_account
      'Allow records to be Deleted for different sales_ledger_account. Prompt user first.
      Dim vCC As New CreditCustomer
      Dim vCC2 As New CreditCustomer
      Dim vRecordSet As CDBRecordSet
      Dim vRecordSet2 As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vError As Boolean

      vCC.Init(pEnv)
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & vCC.GetRecordSetFields(CreditCustomer.CreditCustomerRecordSetTypes.ccurtAll) & " FROM credit_customers ccu WHERE contact_number = " & pOriginalContact)
      While vRecordSet.Fetch() = True And vError = False
        vCC.InitFromRecordSet(pEnv, vRecordSet, CreditCustomer.CreditCustomerRecordSetTypes.ccurtAll)
        vRecordSet2 = pEnv.Connection.GetRecordSet("SELECT " & vCC.GetRecordSetFields(CreditCustomer.CreditCustomerRecordSetTypes.ccurtAll) & " FROM credit_customers ccu WHERE contact_number = " & pDuplicateContact & " AND company = '" & vCC.Company & "'")
        If vRecordSet2.Fetch() = True Then
          vRecordSet2.CloseRecordSet()
          vRecordSet.CloseRecordSet()
          CreditCustomerCheck = True
          Exit Function
        End If
        vRecordSet2.CloseRecordSet()
      End While
      vRecordSet.CloseRecordSet()
      CreditCustomerCheck = False
    End Function

    Public Sub AmalgamateOrganisation(ByVal pEnv As CDBEnvironment, ByVal pJob As JobSchedule, ByVal pParams As CDBParameters)
      Dim vPContact As New Contact(pEnv)
      Dim vDContact As New Contact(pEnv)
      Dim vFields As New CDBFields

      Dim vPrimaryOrganisationNumber As Integer
      Dim vDuplicateOrganisationNumber As Integer

      vPrimaryOrganisationNumber = pParams("OrganisationNumber").IntegerValue
      vDuplicateOrganisationNumber = pParams("AmalgamateOrganisationNumber").IntegerValue

      If vPrimaryOrganisationNumber = vDuplicateOrganisationNumber Then
        RaiseError(DataAccessErrors.daeAmalgamateSameOrganisation) 'You may not amalgamate an organisation with itself
      Else
        If pEnv.Connection.GetCount("branches", Nothing, "organisation_number IN (" & vPrimaryOrganisationNumber & "," & vDuplicateOrganisationNumber & ")") > 1 Then
          RaiseError(DataAccessErrors.daeMergeBothBranches) 'Both organisations are branches
        ElseIf pEnv.Connection.GetCount("duplicate_contacts", Nothing, "contact_number_1 IN(" & vPrimaryOrganisationNumber & "," & vDuplicateOrganisationNumber & ")") > 0 OrElse pEnv.Connection.GetCount("duplicate_contacts", Nothing, "contact_number_2 IN(" & vPrimaryOrganisationNumber & "," & vDuplicateOrganisationNumber & ")") > 0 Then
          RaiseError(DataAccessErrors.daeMergeOrgAlreadyset) 'One of these Organisations has already been set for DeDuplication
        Else
          vPContact.Init(vPrimaryOrganisationNumber)
          vDContact.Init(vDuplicateOrganisationNumber)
          If vPContact.ContactNumber > 0 And vDContact.ContactNumber > 0 Then
            CheckOwnershipValid(pEnv, vPContact, vDContact)
            If pParams.OptionalValue("Confirm", "N") = "N" Then RaiseError(DataAccessErrors.daeAmalgamateOrganisationConfirm)
            vPContact.DoOrgAmalgamation(pJob, pEnv.Connection, vDContact, pParams("Notes").Value, pParams.ParameterExists("Status").Value, pParams.ParameterExists("OwnershipGroup").Value)
          End If
        End If
      End If
    End Sub
  End Class
End Namespace

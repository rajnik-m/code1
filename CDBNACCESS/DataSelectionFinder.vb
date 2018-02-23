Imports System.Data

Namespace Access
  Partial Public Class DataSelection

    Private Enum JoinFlags
      jfJoinToRoles = 1
      jfJoinToSuppressions = 2
      jfJoinToCategories = 4
      jfJoinToCommunications = 8
    End Enum
    Private mvMaxItemsInOracleInClause As Integer = 1000
    Private mvCustomFinders As Dictionary(Of Integer, CustomFinder) 'BR19285 - Changed to Dictionay from Collection
    Private mvSelectItems As New CDBFields

    Private Sub GetFinder(ByVal pDataTable As CDBDataTable)
      'If mvType = DataSelectionTypes.dstActionFinder Then
      '  'If the associated Display List has any of the Related Contact columns as selected items then populate those columns of the supplied DataTable
      '  'Only ContactName, Name & Position are checked because ContactNumber, OrganisationNumber & PhoneNumber will have automatically been added to the Selected Items list because they contain the word "number"
      '  If InStr(mvSelectColumns, "ContactName") > 0 Or InStr(mvSelectColumns, "Name") > 0 Or InStr(mvSelectColumns, "Position") > 0 Then mvDataFinder.Add(ActionRelatedContactData = True)
      'End If
      Dim vHideFields As Boolean
      If mvParameters.Exists("Timeout") Then
        Dim vTimeout As Integer = mvParameters("Timeout").LongValue
        If vTimeout > 0 Then pDataTable.Timeout = vTimeout
      End If

      AddSelectItem(mvParameters)

      Select Case mvType
        Case DataSelectionTypes.dstContactFinder
          ' If pDataTable is empty contact has not been found
          Dim vAttrs As New StringBuilder
          vAttrs.Append("CONTACT_NAME,CONTACT_TELEPHONE,position,position_location,name,o.organisation_number,status_desc,rgb_value")
          If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then vAttrs.Append(",ownership_group_desc,department_desc,ownership_access_level,ownership_access_level_desc")
          Dim vSelectedAttrs As String = GetSelectedAttributes.Replace("c.contact_number", "DISTINCT_CONTACT_NUMBER")
          pDataTable.FillFromSQL(mvEnv, SelectionSQL, vSelectedAttrs, vAttrs.ToString)
          If pDataTable.Rows.Count = 0 AndAlso mvParameters.Exists("ContactNumber") = True Then
            'No records found so check for merged Contact BR17031
            Dim vContactNo As Integer = mvParameters("ContactNumber").LongValue
            Dim vMergedContactNo As Integer = CheckMergedContactOrOrganisation(vContactNo)
            Dim vContact As New Contact(mvEnv)
            vContact.InitWithPrimaryKey(New CDBFields(New CDBField("contact_number", vMergedContactNo))) 'Use contact class to see if the contact number is for a contact or Organisation
            If vContact.ContactType <> Contact.ContactTypes.ctcOrganisation Then
              'Merged contact number is for a contact
              If vMergedContactNo <> vContactNo Then
                'Contact has been merged so reselect using the merged contact number
                mvParameters("ContactNumber").Value = vMergedContactNo.ToString
                pDataTable.FillFromSQL(mvEnv, SelectionSQL, vSelectedAttrs, vAttrs.ToString)
                If pDataTable.Rows.Count = 0 Then 'BR18231
                  'Merged contact not found so return error
                  RaiseError(DataAccessErrors.daeMergedContactNotFound, CStr(vContactNo), CStr(vMergedContactNo))
                End If
              End If
            End If
          End If
          pDataTable.RemoveDuplicateRows("ContactNumber") 'BR18231
          vHideFields = True
          GetCustomMergeData(pDataTable)

        Case DataSelectionTypes.dstOrganisationFinder
          ' If pDataTable is empty Organisation has not been found
          Dim vAttrs As String = "status_desc"
          If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then vAttrs = vAttrs & ",ownership_group_desc,department_desc,ownership_access_level,ownership_access_level_desc"
          Dim vSelectedAttrs As String = GetSelectedAttributes.Replace("o.organisation_number", "DISTINCT_ORGANISATION_NUMBER")
          pDataTable.FillFromSQL(mvEnv, SelectionSQL, vSelectedAttrs, vAttrs)
          If pDataTable.Rows.Count = 0 AndAlso mvParameters.Exists("OrganisationNumber") = True Then
            'No records found so check for merged Organisation BR17031
            Dim vOrganisationNo As Integer = mvParameters("OrganisationNumber").LongValue
            Dim vContactType As Contact.ContactTypes = Contact.ContactTypes.ctcOrganisation
            Dim vMergedOrganisationNo As Integer = CheckMergedContactOrOrganisation(vOrganisationNo)
            Dim vContact As New Contact(mvEnv)
            vContact.InitWithPrimaryKey(New CDBFields(New CDBField("contact_number", vMergedOrganisationNo))) ' Use contact class to see if the contact number is for a contact or Organisation
            If vContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
              'Merged contact number is for an organisation 
              If vMergedOrganisationNo <> vOrganisationNo Then
                'Organisation has been merged so reselect using the merged organisation number
                mvParameters("OrganisationNumber").Value = vMergedOrganisationNo.ToString
                pDataTable.FillFromSQL(mvEnv, SelectionSQL, vSelectedAttrs, vAttrs)
                If pDataTable.Rows.Count = 0 Then 'BR18231
                  'Merged organisation not found so return error
                  RaiseError(DataAccessErrors.daeMergedContactNotFound, CStr(vOrganisationNo), CStr(vMergedOrganisationNo))
                End If
              End If
            End If
          End If
          pDataTable.RemoveDuplicateRows("OrganisationNumber")  'BR18231
          vHideFields = True
          GetCustomMergeData(pDataTable)

        Case DataSelectionTypes.dstContactMailingDocumentsFinder
          Dim vAdditionalItems As String = ""
          If mvParameters.ParameterExists("Fulfilled").Bool = False Then vAdditionalItems = ",,,,"
          pDataTable.FillFromSQL(mvEnv, SelectionSQL, GetSelectedAttributes, vAdditionalItems)

        Case DataSelectionTypes.dstMailingFinder,
             DataSelectionTypes.dstEventPersonnelFinder,
             DataSelectionTypes.dstEventPersonnelAppointmentFinder,
             DataSelectionTypes.dstExamPersonnelFinder
          pDataTable.FillFromSQL(mvEnv, SelectionSQL, GetSelectedAttributes, "")

        Case DataSelectionTypes.dstTextSearch
          ProcessTextSearch(pDataTable)

        Case DataSelectionTypes.dstActionFinder
          mvDataFinder.SetSelectItems(mvParameters)
          'If the associated Display List has any of the Related Contact columns as selected items then populate those columns of the supplied DataTable
          'Only ContactName, Name & Position are checked because ContactNumber, OrganisationNumber & PhoneNumber will have automatically been added to the Selected Items list because they contain the word "number"
          If InStr(mvSelectColumns, "ContactName") > 0 Or InStr(mvSelectColumns, "Name") > 0 Or InStr(mvSelectColumns, "Position") > 0 Then mvDataFinder.AddActionRelatedContactData = True
          mvDataFinder.GetActionsData(pDataTable)
          If mvDataFinder.CustomColumns.Length > 0 Then
            mvSelectColumns = mvSelectColumns & "," & mvDataFinder.CustomColumns
            mvHeadings = mvHeadings & "," & mvDataFinder.CustomHeadings
            Dim vItems() As String = mvDataFinder.CustomColumns.Split(","c)
            For Each vItem As String In vItems
              mvWidths = mvWidths & ",300"
            Next
          End If

        Case DataSelectionTypes.dstEventFinder
          mvDataFinder.SetSelectItems(mvParameters)
          Dim vSQLStatement As New SQLStatement(mvEnv.Connection, mvDataFinder.SelectionSQL)
          pDataTable.FillFromSQL(mvEnv, vSQLStatement)

      End Select
      If vHideFields Then pDataTable.SuppressData()
    End Sub

    Private Function GetSelectedAttributes() As String
      Dim vSelectedAttrs As String = ""
      Select Case mvType
        Case DataSelectionTypes.dstContactFinder
          Dim vContact As New Contact(mvEnv)
          Dim vAddress As New Address(mvEnv)
          vSelectedAttrs = vContact.GetRecordSetFieldsNamePhone & ",date_of_birth,c.department,c.status,c.ownership_group," & vAddress.GetRecordSetFieldsDetails
        Case DataSelectionTypes.dstOrganisationFinder
          Dim vOrg As New Organisation(mvEnv)
          Dim vAddress As New Address(mvEnv)
          vSelectedAttrs = vOrg.GetRecordSetFieldsNamePhone & ",o.status,o.ownership_group," & vAddress.GetRecordSetFieldsDetails
        Case DataSelectionTypes.dstContactMailingDocumentsFinder
          vSelectedAttrs = "mailing_document_number,cmd.mailing_template,mailing_template_desc,label_name,created_by,created_on,cmd.mailing,mailing_desc,earliest_fulfilment_date"
          If mvParameters.ParameterExists("Fulfilled").Bool = True Then vSelectedAttrs = vSelectedAttrs & ",cmd.fulfillment_number,fulfilled_by,fulfilled_on,number_of_documents"
        Case DataSelectionTypes.dstMailingFinder
          vSelectedAttrs = "m.mailing,mailing_desc,mailing_date,mailing_number,mailing_by,number_in_mailing,number_of_emails,number_processed,number_failed,issue_id,number_emails_bounced,number_emails_clicked,number_emails_opened,mh.topic,topic_desc,mh.sub_topic,sub_topic_desc,mh.subject,email_job_number"
        Case DataSelectionTypes.dstEventPersonnelFinder
          vSelectedAttrs = "contact_number,address_number,surname,initials"
        Case DataSelectionTypes.dstEventPersonnelAppointmentFinder
          vSelectedAttrs = "contact_number,start_date,end_date"
        Case DataSelectionTypes.dstExamPersonnelFinder
          vSelectedAttrs = "exam_personnel_id,contact_number,address_number,surname,forenames,initials,valid_from,valid_to,exam_personnel_type,exam_personnel_type_desc"
        Case DataSelectionTypes.dstTextSearch
          vSelectedAttrs = "rank_number,contact_number,event_number,document_number,id_number,id_desc,full_text"
      End Select
      Return vSelectedAttrs
    End Function

    Private Function SelectionSQL() As SQLStatement
      If mvParameters.Count < 1 Then RaiseError(DataAccessErrors.daeNoSelectionData)
      Select Case mvType
        Case DataSelectionTypes.dstContactFinder
          Return GetContactSelectionSQL(True)
        Case DataSelectionTypes.dstOrganisationFinder
          Return GetContactSelectionSQL(False)
        Case DataSelectionTypes.dstContactMailingDocumentsFinder
          Return GetContactMailingDocumentsSelectionSQL()
        Case DataSelectionTypes.dstMailingFinder
          Return GetMailingSelectionSQL()
        Case DataSelectionTypes.dstEventPersonnelFinder
          Return GetEventPersonnelSelectionSQL()
        Case DataSelectionTypes.dstEventPersonnelAppointmentFinder
          Return GetEventPersonnelAppointmentSelectionSQL()
        Case DataSelectionTypes.dstExamPersonnelFinder
          Return GetExamPersonnelSelectionSQL()
        Case DataSelectionTypes.dstTextSearch
          Return GetTextSearchSelectionSQL()
        Case Else
          Return Nothing
      End Select
    End Function

    Private Function GetContactSelectionSQL(ByVal pContact As Boolean) As SQLStatement
      Dim vUseSearchNames As Boolean
      Dim vJoinToSearchNames As Boolean
      Dim vUniSearch As Boolean
      Dim vDefaultAddressOnly As Boolean = False
      If mvEnv.GetConfigOption("cd_advanced_name_searching", False) Then
        vUseSearchNames = True
        If mvEnv.Connection.IsCaseSensitive OrElse
           (mvParameters.Exists("UseSoundex") AndAlso mvParameters("UseSoundex").Bool) OrElse
           (mvParameters.Exists("UseSearchNames") AndAlso mvParameters("UseSearchNames").Bool) Then
          vJoinToSearchNames = True
        End If
      End If
      Dim vSQLStatement As SQLStatement
      Dim vOrderBy As String
      Dim vExternalNumbers As String = GetExternalNumbers()

      Dim vAllAddresses As Boolean

      If mvEnv.GetConfig("uniserv_mail").Length > 0 Then
        Dim vName As String = ""
        Dim vCustFinderExtNumbers As String = vExternalNumbers
        vExternalNumbers = ""
        If pContact Then vName = "" Else vName = mvParameters.OptionalValue("Name", "")
        If mvParameters.HasValue("Name") OrElse mvParameters.HasValue("PreferredForename") OrElse mvParameters.HasValue("Surname") OrElse mvParameters.HasValue("Address") OrElse mvParameters.HasValue("Town") OrElse mvParameters.HasValue("Postcode") Then
          Dim vErrorNumber As Integer = mvEnv.UniservInterface.FindContact(mvParameters.OptionalValue("PreferredForename", ""), mvParameters.OptionalValue("Surname", ""), vName, mvParameters.OptionalValue("BuildingNumber", ""), mvParameters.OptionalValue("Address", ""), mvParameters.OptionalValue("Town", ""), mvParameters.OptionalValue("Postcode", ""), mvParameters.OptionalValue("Country", ""), mvParameters.OptionalValue("DateOfBirth", ""), vExternalNumbers)
          If vExternalNumbers.Length > 0 Then
            vUniSearch = True
            vAllAddresses = True
            If vCustFinderExtNumbers.Length > 0 Then  'De-dup the numbers found by Custom Finder and Uniserv
              Dim vCustFinderArray() As String = vCustFinderExtNumbers.Split(","c)
              Dim vTempExtNumbers As New StringBuilder
              For vIndex As Integer = 0 To vCustFinderArray.Length - 1
                If ("," & vExternalNumbers & ",").Contains("," & vCustFinderArray(vIndex) & ",") Then
                  If vTempExtNumbers.Length > 0 Then vTempExtNumbers.Append(",")
                  vTempExtNumbers.Append(vCustFinderArray(vIndex))
                End If
              Next
              If vTempExtNumbers.Length = 0 Then
                vExternalNumbers = "" 'No record matched in the results of Custom Finder and Uniserv
              Else
                vExternalNumbers = vTempExtNumbers.ToString
              End If
            End If
          End If
        Else
          vExternalNumbers = vCustFinderExtNumbers  'Reset external numbers for any Custom Finder result
        End If
      End If

      Dim vAttrList As String = GetSelectedAttributes()

      If mvParameters.Exists("Country") OrElse mvParameters.Exists("Postcode") OrElse mvParameters.Exists("Town") OrElse
         mvParameters.Exists("Branch") OrElse mvParameters.Exists("Address") OrElse mvParameters.Exists("GovernmentRegion") OrElse
         mvParameters.Exists("BuildingNumber") Then
        vAllAddresses = True
        If mvParameters.HasValue("Address") AndAlso Not mvParameters("Address").Value.Contains("*") Then mvParameters("Address").Value &= "*"

        If mvParameters.Exists("DefaultAddressOnly") AndAlso mvParameters("DefaultAddressOnly").Bool Then
          vDefaultAddressOnly = True
        End If
      End If

      Dim vAnsiJoins As New AnsiJoins()
      Dim vWhereFields As New CDBFields()
      Dim vFirstTable As String
      Dim vWild As String = "*"

      If pContact Then
        vFirstTable = "contacts c"
        If mvParameters.HasValue("Surname") AndAlso Not vUniSearch Then
          If vUseSearchNames Then
            If vJoinToSearchNames Then
              vAnsiJoins.Add("contact_search_names csn", "csn.contact_number", "c.contact_number")
              If mvParameters.Exists("UseSoundex") AndAlso mvParameters("UseSoundex").Bool Then
                vWhereFields.Add("soundex_code", GetSoundexCode(mvParameters("Surname").Value))
              Else
                vWhereFields.Add("search_name", mvParameters("Surname").Value.ToLower, CDBField.FieldWhereOperators.fwoLikeOrEqual)
              End If
              If mvParameters.Exists("UseSearchNames") Then
                If Not mvParameters("UseSearchNames").Bool Then vWhereFields.Add("csn.is_active", "Y")
              Else
                vWhereFields.Add("csn.is_active", "Y")
              End If
            Else
              vWhereFields.Add("surname", mvParameters("Surname").Value, CDBField.FieldWhereOperators.fwoLike)
            End If
          Else
            vWhereFields.Add("surname", mvParameters("Surname").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          End If
        End If

        If mvParameters.Exists("PostPoint") Then
          vAnsiJoins.Add("post_point_recipients ppr", "ppr.contact_number", "c.contact_number")
          vWhereFields.Add("ppr.post_point", mvParameters("PostPoint").Value)
        End If

        If mvParameters.Exists("ExternalReference") Then
          vAnsiJoins.Add("contact_external_links cel", "cel.contact_number", "c.contact_number")
          vWhereFields.Add("cel.external_reference", mvParameters("ExternalReference").Value)
          AddWhereFieldFromParameter(vWhereFields, "DataSource", "data_source")
        End If

        If mvParameters.Exists("PhoneNumber") Then
          vAnsiJoins.Add("communications co", "c.contact_number", "co.contact_number")
          'if user has entered wildcard then don't add another one on the end
          If mvParameters("PhoneNumber").Value.Contains("*") OrElse mvParameters("PhoneNumber").Value.Contains("%") Then vWild = ""
          vWhereFields.Add("co.cli_number", ExtractNumberAndWildcard(mvParameters("PhoneNumber").Value) & vWild, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        End If

        If mvParameters.Exists("EMailAddress") Then
          vAnsiJoins.Add("communications em", "c.contact_number", "em.contact_number")
          vWhereFields.Add("em.number", mvParameters("EMailAddress").Value & "*", CDBField.FieldWhereOperators.fwoLike).SpecialColumn = True
        End If

        If mvParameters.HasValue("RestrictNonHistoricActivity") AndAlso mvParameters("RestrictNonHistoricActivity").Value = "Y" Then
          vWhereFields.Add("cc.valid_from", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoLessThanEqual)
          vWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If

        AddRoleSuppressionActivities(pContact, vAnsiJoins, vWhereFields)

        If vAllAddresses And Not vDefaultAddressOnly Then
          vAnsiJoins.Add("contact_addresses ca", "ca.contact_number", "c.contact_number")
          vAnsiJoins.Add("addresses a", "ca.address_number", "a.address_number")
          vAttrList = vAttrList & ",ca.historical"
        Else
          vAnsiJoins.Add("addresses a", "c.address_number", "a.address_number")
        End If

        If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
          mvEnv.User.AddOwnershipJoins(vAnsiJoins, "c")
          AddWhereFieldFromParameter(vWhereFields, "PrincipalDepartment", "principal_department")
          If mvParameters.Exists("PrincipalUser") Then
            vAnsiJoins.Add("principal_users pu", "pu.contact_number", "c.contact_number")
            vWhereFields.Add("principal_user", mvParameters("PrincipalUser").Value)
          End If
          vAttrList = vAttrList & ",og.ownership_group_desc,d.department_desc,oal.ownership_access_level,oal.ownership_access_level_desc"
        End If

        Dim vType As AnsiJoin.AnsiJoinTypes = AnsiJoin.AnsiJoinTypes.LeftOuterJoin
        Dim vOrganisationParameters As Boolean = mvParameters.Exists("Name") OrElse mvParameters.Exists("OrganisationNumber")
        If vOrganisationParameters Then vType = AnsiJoin.AnsiJoinTypes.InnerJoin

        Dim vFindFromPosition As Boolean = (mvParameters.Exists("ContactPositionNumber"))
        If vFindFromPosition Then vType = AnsiJoin.AnsiJoinTypes.InnerJoin

        vAnsiJoins.Add("contact_positions cp", "c.contact_number", "cp.contact_number", vType)
        vAnsiJoins.Add("organisations o", "cp.organisation_number", "o.organisation_number", vType)
        If vOrganisationParameters Then
          AddWhereFieldFromIntegerParameter(vWhereFields, "OrganisationNumber", "o.organisation_number")
          AddWhereFieldFromParameter(vWhereFields, "Name", "o.name")
          vWhereFields.Add("current", "Y").SpecialColumn = True
        End If
        If vFindFromPosition Then AddWhereFieldFromIntegerParameter(vWhereFields, "ContactPositionNumber", "cp.contact_position_number")
        vAnsiJoins.AddLeftOuterJoin("statuses s", "c.status", "s.status")

        AddWhereFieldFromIntegerParameter(vWhereFields, "ContactNumber", "c.contact_number")
        If Not vUniSearch Then AddWhereFieldFromParameter(vWhereFields, "PreferredForename", "preferred_forename")
        If Not vUniSearch Then AddWhereFieldFromParameter(vWhereFields, "Forenames", "forenames")
        AddWhereFieldFromParameter(vWhereFields, "Title", "title")
        AddWhereFieldFromParameter(vWhereFields, "Initials", "initials")
        AddAddressParameters(vWhereFields, vUniSearch)
        AddWhereFieldFromParameter(vWhereFields, "DiallingCode", "c.dialling_code")
        AddWhereFieldFromParameter(vWhereFields, "StdCode", "c.std_code")
        AddWhereFieldFromParameter(vWhereFields, "Telephone", "c.telephone")
        AddWhereFieldFromDateParameter(vWhereFields, "DateOfBirth", "date_of_birth")
        AddWhereFieldFromParameter(vWhereFields, "Department", "c.department")
        AddWhereFieldFromParameter(vWhereFields, "Status", "c.status")
        AddWhereFieldFromParameter(vWhereFields, "Source", "c.source")
        AddWhereFieldFromParameter(vWhereFields, "OwnershipGroup", "c.ownership_group")

        If vExternalNumbers.Length > 0 Then
          If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle AndAlso vExternalNumbers.AsIList.Count() > mvMaxItemsInOracleInClause Then
            ' BR18151 avoid ORA-01795: maximum number of expressions in a list is 1000 error. BR19632 Apply only to Oracle
            RaiseError(DataAccessErrors.daeTooManyContactsForOracleInClause)
          End If
          vWhereFields.Add("c.contact_number#2", CDBField.FieldTypes.cftInteger, vExternalNumbers, CDBField.FieldWhereOperators.fwoIn)
        End If

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbStatusReasons) AndAlso mvParameters.ParameterExists("StatusReason").Value.Length > 0 Then
          vWhereFields.Add("cast(c.status_reason as varchar(4000))", mvParameters("StatusReason").Value)
        End If

        If vWhereFields.Count = 0 Then RaiseError(DataAccessErrors.daeNoSelectionData)

        Dim vGroupCode As String
        If mvParameters.HasValue("ContactGroup") Then
          vGroupCode = mvParameters("ContactGroup").Value
        Else
          vGroupCode = ContactGroup.DefaultGroupCode
        End If
        If vGroupCode = ContactGroup.DefaultGroupCode Then
          vWhereFields.Add("c.contact_group", vGroupCode, CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("c.contact_group#1", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        Else
          vWhereFields.Add("c.contact_group", vGroupCode)
        End If

        vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)

        If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
          mvEnv.User.AddOwnershipWhere(vWhereFields, "c")
        End If
        vOrderBy = "surname, initials, c.contact_number, CASE WHEN a.address_number = c.address_number THEN 0 ELSE 1 END"
        If vAllAddresses AndAlso Not vDefaultAddressOnly Then vOrderBy &= ", historical"
        vOrderBy &= "," & mvEnv.Connection.DBIsNull("""current""", "'Y'") & "DESC"
        vAttrList &= ",position,position_location,name,o.organisation_number,status_desc"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataRgbValueForStatus) Then
          vAttrList = vAttrList & ",s.rgb_value"
        End If

      Else  '--------------------------------------------------------------------------------------------------------------------
        vFirstTable = "organisations o"

        If mvParameters.HasValue("Name") AndAlso Not vUniSearch Then
          If vUseSearchNames Then
            If vJoinToSearchNames Then
              vAnsiJoins.Add("contact_search_names csn", "csn.contact_number", "o.organisation_number")
              If mvParameters.Exists("UseSoundex") AndAlso mvParameters("UseSoundex").Bool Then
                vWhereFields.Add("soundex_code", GetSoundexCode(mvParameters("Name").Value))
              Else
                vWhereFields.Add("search_name", mvParameters("Name").Value.ToLower, CDBField.FieldWhereOperators.fwoLikeOrEqual)
              End If
              If mvParameters.Exists("UseSearchNames") Then
                If Not mvParameters("UseSearchNames").Bool Then vWhereFields.Add("csn.is_active", "Y")
              Else
                vWhereFields.Add("csn.is_active", "Y")
              End If
            Else
              vWhereFields.Add("Name", mvParameters("Name").Value, CDBField.FieldWhereOperators.fwoLike)
            End If
          Else
            vWhereFields.Add("Name", mvParameters("Name").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
          End If
        End If

        If mvParameters.Exists("ExternalReference") Then
          vAnsiJoins.Add("contact_external_links cel", "cel.contact_number", "o.organisation_number")
          vWhereFields.Add("cel.external_reference", mvParameters("ExternalReference").Value)
          AddWhereFieldFromParameter(vWhereFields, "DataSource", "data_source")
        End If

        If mvParameters.HasValue("RestrictNonHistoricActivity") AndAlso mvParameters("RestrictNonHistoricActivity").Value = "Y" Then
          vWhereFields.Add("cc.valid_from", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoLessThanEqual)
          vWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, TodaysDate(), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If

        AddRoleSuppressionActivities(pContact, vAnsiJoins, vWhereFields)

        If mvParameters.Exists("EMailAddress") OrElse mvParameters.Exists("PhoneNumber") OrElse mvParameters.Exists("WebAddress") Then vAllAddresses = True

        If vAllAddresses And Not vDefaultAddressOnly Then
          vAnsiJoins.Add("organisation_addresses ca", "ca.organisation_number", "o.organisation_number")
          vAnsiJoins.Add("addresses a", "ca.address_number", "a.address_number")
          vAttrList = vAttrList & ",ca.historical"
        Else
          vAnsiJoins.Add("addresses a", "o.address_number", "a.address_number")
        End If

        If mvParameters.Exists("PhoneNumber") Then
          vAnsiJoins.Add("communications co", "a.address_number", "co.address_number")
          If mvParameters.OptionalValue("UseContactRestriction", "N") = "Y" Then
            vWhereFields.Add("co.contact_number", CDBField.FieldTypes.cftInteger)     'Jira 663 - removed code to make finder consistent with rich client
          End If
          vAllAddresses = True
          vWild = "*"  'if user has entered wildcard then don't add another one on the end
          If mvParameters("PhoneNumber").Value.Contains("*") OrElse mvParameters("PhoneNumber").Value.Contains("%") Then vWild = ""
          vWhereFields.Add("co.cli_number", ExtractNumberAndWildcard(mvParameters("PhoneNumber").Value) & vWild, CDBField.FieldWhereOperators.fwoLikeOrEqual)
        End If
        If mvParameters.Exists("EMailAddress") Then
          vAnsiJoins.Add("communications em", "a.address_number", "em.address_number")
          If mvParameters.OptionalValue("UseContactRestriction", "N") = "Y" Then
            vWhereFields.Add("em.contact_number", CDBField.FieldTypes.cftInteger)     'Jira 663 - removed code to make finder consistent with rich client
          End If
          vAllAddresses = True
          vWhereFields.Add("em.number", mvParameters("EMailAddress").Value & "*", CDBField.FieldWhereOperators.fwoLike).SpecialColumn = True
        End If
        If mvParameters.Exists("WebAddress") Then
          vAnsiJoins.Add("communications wa", "a.address_number", "wa.address_number")
          If mvParameters.OptionalValue("UseContactRestriction", "N") = "Y" Then
            vWhereFields.Add("wa.contact_number", CDBField.FieldTypes.cftInteger)     'Jira 663 - removed code to make finder consistent with rich client
          End If
          vAllAddresses = True
          If vWhereFields.ContainsKey("wa.number") Then
            vWhereFields.Add("wa.number#2", mvParameters("WebAddress").Value & "*", CDBField.FieldWhereOperators.fwoLike).SpecialColumn = True
          Else
            vWhereFields.Add("wa.number", mvParameters("WebAddress").Value & "*", CDBField.FieldWhereOperators.fwoLike).SpecialColumn = True
          End If
        End If

        If mvParameters.Exists("NoMemberOrganisations") AndAlso mvParameters("NoMemberOrganisations").Value = "Y" Then
          vAnsiJoins.Add("members mem", "o.organisation_number", "mem.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
          vWhereFields.Add("mem.cancellation_reason", "", CDBField.FieldWhereOperators.fwoNullOrEqual)
          vWhereFields.Add("mem.member_number", "", CDBField.FieldWhereOperators.fwoNullOrEqual)
        End If

        If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
          mvEnv.User.AddOwnershipJoins(vAnsiJoins, "o")
          If mvParameters.Exists("PrincipalUser") Then
            vAnsiJoins.Add("principal_users pu", "pu.contact_number", "o.organisation_number")
            vWhereFields.Add("principal_user", mvParameters("PrincipalUser").Value)
          End If
          vAttrList = vAttrList & ",og.ownership_group_desc,d.department_desc,oal.ownership_access_level,oal.ownership_access_level_desc"
        End If
        vAnsiJoins.AddLeftOuterJoin("statuses s", "o.status", "s.status")

        AddWhereFieldFromIntegerParameter(vWhereFields, "OrganisationNumber", "o.organisation_number")
        AddWhereFieldFromParameter(vWhereFields, "Abbreviation", "o.abbreviation")
        AddAddressParameters(vWhereFields, vUniSearch)
        AddWhereFieldFromParameter(vWhereFields, "DiallingCode", "o.dialling_code")
        AddWhereFieldFromParameter(vWhereFields, "StdCode", "o.std_code")
        AddWhereFieldFromParameter(vWhereFields, "Telephone", "o.telephone")
        AddWhereFieldFromParameter(vWhereFields, "Department", "o.department")
        AddWhereFieldFromParameter(vWhereFields, "Status", "o.status")
        AddWhereFieldFromParameter(vWhereFields, "Source", "o.source")
        AddWhereFieldFromParameter(vWhereFields, "OwnershipGroup", "o.ownership_group")

        If vExternalNumbers.Length > 0 Then
          If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle AndAlso vExternalNumbers.AsIList.Count() > mvMaxItemsInOracleInClause Then
            ' BR18151 avoid ORA-01795: maximum number of expressions in a list is 1000 error. BR19632 Apply only to Oracle
            RaiseError(DataAccessErrors.daeTooManyOrganisationsForOracleInClause)
          End If
          vWhereFields.Add("o.organisation_number#2", CDBField.FieldTypes.cftInteger, vExternalNumbers, CDBField.FieldWhereOperators.fwoIn)
        End If

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbStatusReasons) AndAlso mvParameters.Exists("StatusReason") AndAlso mvParameters("StatusReason").Value.Length > 0 Then
          vWhereFields.Add("cast(o.status_reason as varchar(4000))", mvParameters("StatusReason").Value)
        End If

        If Not mvParameters.Exists("PortalOrgSearch") OrElse Not mvParameters("PortalOrgSearch").Bool Then
          If vWhereFields.Count = 0 Then RaiseError(DataAccessErrors.daeNoSelectionData)
        End If

        Dim vGroupCode As String
        If mvParameters.HasValue("OrganisationGroup") Then
          vGroupCode = mvParameters("OrganisationGroup").Value
        Else
          vGroupCode = OrganisationGroup.DefaultGroupCode
        End If
        If vGroupCode = OrganisationGroup.DefaultGroupCode Then
          vWhereFields.Add("o.organisation_group", vGroupCode, CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("o.organisation_group#1", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        Else
          vWhereFields.Add("organisation_group", vGroupCode)
        End If

        If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
          mvEnv.User.AddOwnershipWhere(vWhereFields, "o")
          AddWhereFieldFromParameter(vWhereFields, "PrincipalDepartment", "principal_department")
        End If
        vOrderBy = "sort_name, name, o.organisation_number, CASE WHEN a.address_number = o.address_number THEN 0 ELSE 1 END"
        If vAllAddresses AndAlso Not vDefaultAddressOnly Then vOrderBy &= ", historical"
        vAttrList &= ",status_desc"
      End If
      vSQLStatement = New SQLStatement(mvEnv.Connection, vAttrList, vFirstTable, vWhereFields, vOrderBy, vAnsiJoins)
      If mvParameters.Exists("NumberOfRows") Then
        Dim vMaxRows As Integer = mvParameters("NumberOfRows").LongValue + 1
        If vMaxRows > 0 Then vSQLStatement.MaxRows = vMaxRows
      End If
      If mvEnv.GetConfigOption("cd_finders_nolock") Then vSQLStatement.NoLock = True
      Return vSQLStatement
    End Function

    Private Sub AddAddressParameters(ByVal vWhereFields As CDBFields, ByVal pUniSearch As Boolean)
      AddWhereFieldFromParameter(vWhereFields, "Country", "country")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then AddWhereFieldFromParameter(vWhereFields, "BuildingNumber", "building_number")
      If Not pUniSearch Then
        AddWhereFieldFromParameter(vWhereFields, "Postcode", "postcode")
        AddWhereFieldFromParameter(vWhereFields, "Town", "town")
        AddWhereFieldFromParameter(vWhereFields, "GovernmentRegion", "mosaic_code")
        AddWhereFieldFromParameter(vWhereFields, "Address", "address", CDBField.FieldTypes.cftMemo)
      End If
      AddWhereFieldFromParameter(vWhereFields, "Branch", "branch")
      AddWhereFieldFromParameter(vWhereFields, "HouseName", "house_name")
    End Sub

    Private Sub AddRoleSuppressionActivities(ByVal pContact As Boolean, ByVal vAnsiJoins As AnsiJoins, ByVal vWhereFields As CDBFields)
      If mvParameters.Exists("Role") Then
        If pContact Then
          vAnsiJoins.Add("contact_roles cr", "cr.contact_number", "c.contact_number")
        Else
          vAnsiJoins.Add("contact_roles cr", "cr.organisation_number", "o.organisation_number")
        End If
        vWhereFields.Add("cr.role", mvParameters("Role").Value)
      End If

      If mvParameters.Exists("Suppression") Then
        If pContact Then
          vAnsiJoins.Add("contact_suppressions cs", "cs.contact_number", "c.contact_number")
        Else
          vAnsiJoins.Add("organisation_suppressions cs", "cs.organisation_number", "o.organisation_number")
        End If
        vWhereFields.Add("cs.mailing_suppression", mvParameters("Suppression").Value)

        If mvParameters.Exists("SuppressionValidFrom") Then
          vWhereFields.Add("cs.valid_from", CDBField.FieldTypes.cftDate, mvParameters("SuppressionValidFrom").Value, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoGreaterThanEqual)
          vWhereFields.Add("cs.valid_from#2", CDBField.FieldTypes.cftDate, mvParameters("SuppressionValidFrom").Value, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLessThanEqual)
          vWhereFields.Add("cs.valid_to", CDBField.FieldTypes.cftDate, mvParameters("SuppressionValidFrom").Value, CDBField.FieldWhereOperators.fwoGreaterThan Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
        End If
        If mvParameters.Exists("SuppressionValidTo") Then
          vWhereFields.Add("cs.valid_to#2", CDBField.FieldTypes.cftDate, mvParameters("SuppressionValidTo").Value, CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoLessThanEqual)
          vWhereFields.Add("cs.valid_to#3", CDBField.FieldTypes.cftDate, mvParameters("SuppressionValidTo").Value, CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoGreaterThanEqual)
          vWhereFields.Add("cs.valid_from#3", CDBField.FieldTypes.cftDate, mvParameters("SuppressionValidTo").Value, CDBField.FieldWhereOperators.fwoLessThan Or CDBField.FieldWhereOperators.fwoCloseBracketTwice)
        End If
      End If

      If mvParameters.Exists("Activity") OrElse mvParameters.Exists("ActivityValue") Then
        If pContact Then
          vAnsiJoins.Add("contact_categories cc", "cc.contact_number", "c.contact_number")
        Else
          vAnsiJoins.Add("organisation_categories cc", "cc.organisation_number", "o.organisation_number")
        End If
        If mvParameters.Exists("Activity") Then vWhereFields.Add("cc.activity", mvParameters("Activity").Value)
        If mvParameters.Exists("ActivityValue") Then vWhereFields.Add("cc.activity_value", mvParameters("ActivityValue").Value)
        If mvParameters.Exists("ActivityValidTo") Then
          If mvParameters.Exists("ActivityValidFrom") Then
            vWhereFields.Add("cc.valid_from", CDBField.FieldTypes.cftDate, mvParameters("ActivityValidFrom").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
            vWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, mvParameters("ActivityValidTo").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
          Else
            vWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, mvParameters("ActivityValidTo").Value)
          End If
        ElseIf mvParameters.Exists("ActivityValidFrom") Then
          vWhereFields.Add("cc.valid_from", CDBField.FieldTypes.cftDate, mvParameters("ActivityValidFrom").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)
          vWhereFields.Add("cc.valid_to", CDBField.FieldTypes.cftDate, mvParameters("ActivityValidFrom").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If
      End If
    End Sub

    Public Function CheckMergedContact(ByVal pEnv As CDBEnvironment, ByVal pContactNumber As Integer, ByRef pContactType As Contact.ContactTypes, ByRef pFields As CDBFields) As Integer
      'If contact finder does not find any records then see if Contact has been merged
      'This will return the Master number with pFields set with some Contact fields (for DataImport)
      Dim vAttr As String
      Dim vTable As String
      If pContactType = Contact.ContactTypes.ctcOrganisation Then
        vTable = "organisations"
        vAttr = "organisation_number"
      Else
        vTable = "contacts"
        vAttr = "contact_number"
      End If
      Dim vAttrs As String = "duplicate, master, address_number, " & vAttr
      If pContactType <> Contact.ContactTypes.ctcOrganisation Then vAttrs += ", contact_type, label_name, contact_group"
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add(vTable & " x", "d.master", "x." & vAttr)
      Dim vWhereFields As New CDBFields()
      vWhereFields.Add("duplicate", pContactNumber)
      Dim vSQLStatement As New SQLStatement(pEnv.Connection, vAttrs, "dba_notes d", vWhereFields, "", vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet
      If vRS.Fetch Then
        CheckMergedContact = vRS.Fields("master").LongValue
        pFields = New CDBFields   '(DataImport only)
        With pFields
          .Add(vAttr, vRS.Fields(vAttr).LongValue)
          .Add("address_number", vRS.Fields("address_number").LongValue)
          If pContactType <> Contact.ContactTypes.ctcOrganisation Then
            .Add("label_name", vRS.Fields("label_name").Value)
            .Add("contact_group", vRS.Fields("contact_group").Value)
            If vRS.Fields("contact_type").Value.Length > 0 Then
              If vRS.Fields("contact_type").Value = "C" Then
                pContactType = Contact.ContactTypes.ctcContact
              Else
                pContactType = Contact.ContactTypes.ctcJoint
              End If
            End If
          End If
        End With
      End If
      vRS.CloseRecordSet()
    End Function
    ''' <summary>
    ''' Return the current Contact number for a merged contact or organisation.
    ''' BR17031
    ''' </summary>
    ''' <param name="pContactNumber">The merged contact (or organisation) number to search for</param>
    ''' <returns>The current contact number for the merged contact, or pContactNumber if the contact number is not merged.</returns>
    ''' <remarks>Relies on table dba_notes, and does not distinguish between contacts or organisations </remarks>
    Public Function CheckMergedContactOrOrganisation(ByVal pContactNumber As Integer) As Integer
      Dim vUltimateMaster As Integer = FindMasterContact(pContactNumber)
      Return vUltimateMaster
    End Function
    ''' <summary>
    ''' Finds the current contact number for a merged contact 
    ''' BR17031
    ''' </summary>
    ''' <param name="pDuplicateNumber">The merged contact number to look for.</param>
    ''' <returns>The current contact number for the merged contact or the merged contact number if the contact number is not merged</returns>
    ''' <remarks></remarks>
    Private Function FindMasterContact(pDuplicateNumber As Integer) As Integer
      Dim vFields As String = "master,duplicate,merged_on,notes"
      Dim vTable As String = "dba_notes"
      Dim vWhereFields As New CDBFields(New CDBField("duplicate", pDuplicateNumber))
      Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vFields, vTable, vWhereFields, "merged_on DESC")
      Dim vDataTable As DataTable = vSQLStatement.GetDataTable()
      Dim vDataRow As DataRow
      Dim vMasterNumber As Integer
      If vDataTable.Rows.Count > 0 Then
        vDataRow = vDataTable.Rows(0)
        vMasterNumber = CInt(vDataRow.Item("master"))
        Return FindMasterContact(vMasterNumber)
      Else
        Return pDuplicateNumber
      End If
    End Function




    Private Sub GetCustomMergeData(ByVal pDT As CDBDataTable)
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCustomFinderTab) AndAlso mvEnv.GetConfigOption("option_custom_data") Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("client", mvEnv.ClientCode)
        Dim vCustomField As String = ""
        Dim vAttrName As String = ""
        Dim vGroupParam As String = ""
        Select Case mvType
          Case DataSelectionTypes.dstActionFinder
            vWhereFields.Add("usage_code", "A")
            vGroupParam = "contact_group"
            vAttrName = "ActionNumber"
            vCustomField = "action_number"
          Case DataSelectionTypes.dstContactFinder
            vWhereFields.Add("usage_code", "C")
            vGroupParam = "contact_group"
            vAttrName = "ContactNumber"
            vCustomField = "contact_number"
          Case DataSelectionTypes.dstOrganisationFinder
            vWhereFields.Add("usage_code", "O")
            vGroupParam = "organisation_group"
            vAttrName = "OrganisationNumber"
            vCustomField = "contact_number"
            'Case DataSelectionTypes.dstser
            '  vWhereFields.Add("usage_code", cftCharacter, "S")
            '  vGroupParam = "contact_group"
            '  vAttrName = "ContactNumber"
            '  vCustomField = "contact_number"
        End Select
        If mvParameters.Exists(vGroupParam) Then
          vWhereFields.Add("contact_group", mvParameters(vGroupParam).Value)
        End If
        Dim vCMD As New CustomMergeData(mvEnv)
        Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, vCMD.GetRecordSetFields, "custom_merge_data cmd", vWhereFields, "sequence_number").GetRecordSet
        Dim vCMDS As New CollectionList(Of CustomMergeData)
        While vRecordSet.Fetch
          vCMD.InitFromRecordSet(vRecordSet)
          vCMDS.Add(vCMDS.Count.ToString, vCMD)
          vCMD = New CustomMergeData(mvEnv)
        End While
        vRecordSet.CloseRecordSet()

        Dim vCount As Integer
        Dim vItemNo As String
        Dim vCheckRow As CDBDataRow
        Dim vIncludeNo As String
        Dim vCustomColumns As New StringBuilder
        Dim vCustomHeadings As New StringBuilder

        For Each vCMD In vCMDS
          pDT.AddColumnsFromList(vCMD.StandardAttributeNames)
          If vCustomColumns.Length > 0 Then vCustomColumns.Append(",")
          vCustomColumns.Append(vCMD.StandardAttributeNames)
          If vCustomHeadings.Length > 0 Then vCustomHeadings.Append(",")
          vCustomHeadings.Append(vCMD.AttributeCaptions)
          Dim vInclude As New StringBuilder
          For Each vRow As CDBDataRow In pDT.Rows
            vIncludeNo = vRow.Item(vAttrName)
            If vIncludeNo.Length > 0 Then
              If vInclude.Length > 0 Then vInclude.Append(",")
              vInclude.Append(vIncludeNo)
              vCount += 1
              If vCount >= 250 OrElse vRow Is pDT.Rows(pDT.Rows.Count - 1) Then
                If vInclude.Length > 0 Then
                  'now retrieve custom information
                  Dim vRecordSet2 As CDBRecordSet = vCMD.GetRecordSet(vInclude.ToString)
                  While vRecordSet2.Fetch()
                    vItemNo = vRecordSet2.Fields(vCustomField).Value
                    For Each vCheckRow In pDT.Rows
                      If vCheckRow.Item(vAttrName) = vItemNo Then vCMD.SetDataRow(vCheckRow, vRecordSet2, pDT.Columns)
                    Next
                  End While
                  vRecordSet2.CloseRecordSet()
                End If
                vInclude = New StringBuilder
                vCount = 0
              End If
            End If
          Next
        Next

        If vCustomColumns.Length > 0 Then
          mvSelectColumns = mvSelectColumns & "," & vCustomColumns.ToString
          mvHeadings = mvHeadings & "," & vCustomHeadings.ToString
          Dim vItems As String() = vCustomColumns.ToString.Split(","c)
          For vIndex As Integer = 0 To UBound(vItems)
            mvWidths = mvWidths & ",300"
          Next
        End If

      End If
    End Sub

    Private Function GetExternalNumbers() As String
      Dim vExternalNumbers As String = ""
      Dim vCustomFinderControls As CustomFinderControls = GetCustomFinderControls(BaseCDBControls(Nothing))
      If vCustomFinderControls.Count > 0 Then
        For Each vControl As CustomFinderControl In vCustomFinderControls
          If mvSelectItems.Exists(vControl.ParameterName) Then
            If mvCustomFinders.ContainsKey(vControl.CustomFinder) Then 'BR19285 - Error occurs if mvCustomFinder does not contain vControl.CustomFinder
              mvCustomFinders(vControl.CustomFinder).AddRestriction(vControl, mvSelectItems(vControl.ParameterName))
              mvSelectItems.Remove(vControl.ParameterName)       'Make sure it doesn't get used again?
            End If
          End If
        Next
        For Each vKeyValuePair As KeyValuePair(Of Integer, CustomFinder) In mvCustomFinders
          Dim vCustomFinder As CustomFinder = vKeyValuePair.Value
          Dim vNumbers As String = vCustomFinder.GetContactNumbers
          If Len(vNumbers) > 0 Then
            If Len(vExternalNumbers) > 0 Then vExternalNumbers = vExternalNumbers & ","
            vExternalNumbers = vExternalNumbers & vNumbers
          End If
        Next
      End If
      Return vExternalNumbers
    End Function

    Public Function BaseCDBControls(pPageType As String) As CDBControls
      Dim vBaseControls As New CDBControls
      Dim vPage As New EplPage(mvEnv)
      Dim vPageType As String
      If String.IsNullOrEmpty(pPageType) Then
        vPageType = GetFinderPageType()
      Else
        vPageType = pPageType
      End If
      vPage.Init(vPageType, mvGroupCode)
      vBaseControls.GetPageControls(mvEnv, vPageType, vPage.PageNumbers, True)
      Return vBaseControls
    End Function

    Public Function CDBControls() As CDBControls
      Dim vControl As CDBControl
      Dim vTabNumber As Integer

      Dim vBaseControls As CDBControls = BaseCDBControls(Nothing)
      'For contacts and organisations get any custom controls
      If mvType = DataSelectionTypes.dstContactFinder OrElse mvType = DataSelectionTypes.dstOrganisationFinder Then
        Dim vCustomControls As CustomFinderControls = GetCustomFinderControls(vBaseControls)
        'If there are any custom controls then add them in to the correct tab
        If vCustomControls.Count > 0 Then
          Dim vControls As New CDBControls
          For Each vControl In vBaseControls
            If vControl.ControlType = "tab" Then
              vTabNumber = vTabNumber + 1
              AddTabXControls(vControls, vCustomControls, vTabNumber)
            End If
            vControls.Add(vControl)
          Next
          vTabNumber = vTabNumber + 1       'Add any required at the end
          AddTabXControls(vControls, vCustomControls, vTabNumber)
          Return vControls
        Else
          Return vBaseControls
        End If
      Else
        Return vBaseControls
      End If
    End Function

    Private Sub AddTabXControls(ByVal pControls As CDBControls, ByVal pCustomFinderControls As CustomFinderControls, ByRef pTabNumber As Integer)
      Dim vContactGroup As String = ""
      Dim vAdded As Boolean
      Do
        vAdded = False
        For Each vControl As CustomFinderControl In pCustomFinderControls
          If vControl.TabNumber = pTabNumber Then
            If vAdded = False Or vContactGroup <> vControl.ContactGroupCode Then
              Dim vTabControl As CDBControl = New CDBControl(mvEnv)
              vTabControl.InitTabControl(vControl.CustomFinderDesc, vControl.ContactGroupCode)
              pControls.Add(vTabControl)
              vContactGroup = vControl.ContactGroupCode
            End If
            Dim vNewControl As New CDBControl(mvEnv)
            vNewControl.InitFromCustomFinderControl(vControl)
            pControls.Add(vNewControl)
            vAdded = True
          End If
        Next
        If vAdded Then pTabNumber += 1
      Loop While vAdded
    End Sub

    Private Function GetCustomFinderControls(ByVal pBaseControls As CDBControls) As CustomFinderControls
      Dim vCustomFinderControls As New CustomFinderControls

      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCustomFinderTab) AndAlso mvEnv.GetConfigOption("option_custom_data") Then
        mvCustomFinders = New Dictionary(Of Integer, CustomFinder)
        If mvType = DataSelectionTypes.dstContactFinder OrElse mvType = DataSelectionTypes.dstOrganisationFinder Then
          'Need to get all the custom finder controls as there could be multiple custom finders for different groups with the same attributes
          Dim vGroupCode As String = ""
          If mvType = DataSelectionTypes.dstContactFinder AndAlso mvSelectItems.ContainsKey("contact_group") Then
            vGroupCode = mvSelectItems("contact_group").Value
          ElseIf mvType = DataSelectionTypes.dstOrganisationFinder AndAlso mvSelectItems.ContainsKey("organisation_group") Then
            vGroupCode = mvSelectItems("organisation_group").Value
          End If
          Dim vAddAllGroups As Boolean = False
          If mvSelectItems.Count > 0 AndAlso mvSelectItems.ContainsKey("AllGroups") = False Then
            'Trying to find contacts / organisations so always get all groups so that the parameter names get set correctly
            mvSelectItems.Add(New CDBField("AllGroups", "Y"))
            vAddAllGroups = True
          End If
          vCustomFinderControls = New CustomFinderControls
          Dim vRecordSet As CDBRecordSet = GetCustomControlsSQL.GetRecordSet
          Dim vControl As CustomFinderControl
          Dim vLastFinder As Integer
          Dim vCustomFinder As CustomFinder
          Dim vStartCountAt As Integer
          While vRecordSet.Fetch()
            If pBaseControls.Exists(vRecordSet.Fields("attribute_name").Value) _
              OrElse vRecordSet.Fields("attribute_name").Value = "external_reference" _
              OrElse vRecordSet.Fields("attribute_name").Value = "number" Then
              vStartCountAt = 2
            Else
              vStartCountAt = 0
            End If
            vControl = vCustomFinderControls.AddFromRecordSet(mvEnv, vRecordSet, vStartCountAt)
            If vControl.CustomFinder <> vLastFinder Then
              If (vAddAllGroups = True AndAlso vGroupCode.Length > 0 AndAlso vGroupCode = vRecordSet.Fields("contact_group").Value) OrElse (vAddAllGroups = False OrElse vGroupCode.Length = 0) Then
                vCustomFinder = New CustomFinder(mvEnv)
                vCustomFinder.Init(vControl.CustomFinder, vRecordSet.Fields("select_sql").Value)
                mvCustomFinders.Add(vControl.CustomFinder, vCustomFinder)
              End If
              vLastFinder = vControl.CustomFinder
            End If
          End While
          vRecordSet.CloseRecordSet()
          If vAddAllGroups Then mvSelectItems.Remove("AllGroups") 'Don't know what effect leaving this will have so remove the item we added
        End If
      End If
      Return vCustomFinderControls
    End Function

    Private Function GetCustomControlsSQL() As SQLStatement
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("client", mvEnv.ClientCode)
      vWhereFields.Add("tab_number", CDBField.FieldTypes.cftInteger, "", CDBField.FieldWhereOperators.fwoNotEqual)
      Dim vAllGroups As Boolean
      If mvSelectItems.ContainsKey("AllGroups") Then vAllGroups = mvSelectItems("AllGroups").Bool
      If mvType = DataSelectionTypes.dstContactFinder Then
        If mvSelectItems.Exists("contact_group") AndAlso vAllGroups = False Then
          vWhereFields.Add("contact_group", mvSelectItems("contact_group").Value)
        Else
          'vWhereFields.Add()
          vWhereFields.Add("contact_group", "(SELECT contact_group FROM contact_groups WHERE client = '" & mvEnv.ClientCode & "')", CDBField.FieldWhereOperators.fwoIn Or CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("contact_group#2", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        End If
      Else
        If mvSelectItems.Exists("organisation_group") AndAlso vAllGroups = False Then
          vWhereFields.Add("contact_group", mvSelectItems("organisation_group").Value)
        Else
          vWhereFields.Add("contact_group", "(SELECT organisation_group FROM organisation_groups WHERE client = '" & mvEnv.ClientCode & "')", CDBField.FieldWhereOperators.fwoIn Or CDBField.FieldWhereOperators.fwoOpenBracket)
          vWhereFields.Add("contact_group#2", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        End If
      End If
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("custom_finder_controls cfc", "cf.custom_finder", "cfc.custom_finder")
      vAnsiJoins.Add("maintenance_attributes ma", "cfc.table_name", "ma.table_name", "cfc.attribute_name", "ma.attribute_name")
      Dim vCustomFinderControl As New CustomFinderControl(mvEnv)
      Dim vMaintenanceAttribute As New MaintenanceAttribute(mvEnv)
      vMaintenanceAttribute.InitForAdditionalInfo(False)
      Return New SQLStatement(mvEnv.Connection, vCustomFinderControl.GetRecordSetFields & "," & vMaintenanceAttribute.GetRecordSetFields & ",tab_number,custom_finder_desc,tag_width,contact_group,db_name,select_sql", "custom_finders cf", vWhereFields, "cf.tab_number, cf.contact_group, cfc.sequence_number", vAnsiJoins)
    End Function

    Public Function GetFinderPageType() As String
      Dim vResult As String = ""
      Select Case mvType
        Case DataSelectionTypes.dstContactFinder
          If mvEnv.GetConfig("uniserv_mail").Length > 0 Then vResult = "PLCU" Else vResult = "PLCO"
        Case DataSelectionTypes.dstOrganisationFinder
          If mvEnv.GetConfig("uniserv_mail").Length > 0 Then vResult = "PLOU" Else vResult = "PLOR"
        Case DataSelectionTypes.dstFindBatch, DataSelectionTypes.dstFindOpenBatch
          vResult = "PLFB"
        Case DataSelectionTypes.dstFindCCCA
          vResult = "PLFA"
        Case DataSelectionTypes.dstFindCovenant
          vResult = "PLFC"
        Case DataSelectionTypes.dstFindDirectDebit
          vResult = "PLFD"
        Case DataSelectionTypes.dstFindEvent
          vResult = "PLFE"
        Case DataSelectionTypes.dstFindEventBooking
          vResult = "PLEB"
        Case DataSelectionTypes.dstEventPersonnelFinder
          vResult = "PLEP"
        Case DataSelectionTypes.dstFindGAD
          vResult = "PLFG"
        Case DataSelectionTypes.dstFindInvoice
          vResult = "PLFI"
        Case DataSelectionTypes.dstFindLegacy
          vResult = "PLLG"
        Case DataSelectionTypes.dstFindMeeting
          vResult = "PLME"
        Case DataSelectionTypes.dstFindMember
          vResult = "PLFM"
        Case DataSelectionTypes.dstFindPaymentPlan
          vResult = "PLFP"
        Case DataSelectionTypes.dstFindProduct
          vResult = "PLPR"
        Case DataSelectionTypes.dstFindStandingOrder, DataSelectionTypes.dstFindManualSOReconciliation
          vResult = "PLFS"
        Case DataSelectionTypes.dstFindTransaction
          vResult = "PLFT"
        Case DataSelectionTypes.dstFindVenue
          vResult = "PLEV"
        Case DataSelectionTypes.dstFindGiveAsYouEarn
          vResult = "PLGY"
        Case DataSelectionTypes.dstFindPurchaseOrder
          vResult = "PLFO"
        Case DataSelectionTypes.dstFindInternalResource
          vResult = "PLFN"
        Case DataSelectionTypes.dstContactMailingDocumentsFinder
          vResult = "PLMD"
        Case DataSelectionTypes.dstFindServiceProduct
          vResult = "PLSP"
        Case DataSelectionTypes.dstFindCommunication
          vResult = "PLCM"
        Case DataSelectionTypes.dstDistinctDocuments, DataSelectionTypes.dstDistinctExternalDocuments
          vResult = "PLDO"
        Case DataSelectionTypes.dstDuplicateContacts
          vResult = "PLCD"
        Case DataSelectionTypes.dstDuplicateOrganisations
          vResult = "PLOD"
        Case DataSelectionTypes.dstMailingFinder
          vResult = "PLMA"
        Case DataSelectionTypes.dstTextSearch
          vResult = "PLTS"
        Case DataSelectionTypes.dstActionFinder
          vResult = "PLAF"
        Case DataSelectionTypes.dstSelectItemSelectionSets
          vResult = "PLSS"
        Case DataSelectionTypes.dstFindCampaign
          vResult = "PLCP"
        Case DataSelectionTypes.dstFindPostTaxPayrollGiving
          vResult = "PLPG"
        Case DataSelectionTypes.dstAppealCollections
          vResult = "PLAP"
        Case DataSelectionTypes.dstFindStandardDocuments
          vResult = "PLSD"
        Case DataSelectionTypes.dstFundraisingPaymentSchedule
          vResult = "PFPS"
        Case DataSelectionTypes.dstQueryByExampleContacts
          vResult = "QEC*"
        Case DataSelectionTypes.dstQueryByExampleOrganisations
          vResult = "QEO*"
        Case DataSelectionTypes.dstQueryByExampleEvents
          vResult = "QEE*"
        Case DataSelectionTypes.dstDuplicateContactRecords
          vResult = "PDPC"
        Case DataSelectionTypes.dstPostcodeProximity
          vResult = "PLPC"
        Case DataSelectionTypes.dstExamPersonnelFinder
          vResult = "PLXP"
        Case DataSelectionTypes.dstFundraisingRequests
          vResult = "LFFR"
        Case DataSelectionTypes.dstExamScheduleFinder
          vResult = "XSCH"
        Case DataSelectionTypes.dstFindCPDCyclePeriods
          vResult = "PLCE"
        Case DataSelectionTypes.dstFindCPDPoints
          vResult = "PLCN"
      End Select
      Return vResult
    End Function

    Public Sub AddSelectItem(ByVal pParams As CDBParameters, Optional ByVal pFWO As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoEqual)
      Dim vName As String
      Dim vCount As Integer

      For Each vParam As CDBParameter In pParams
        If mvSelectItems.Exists(vParam.Name) Then
          vCount = 1                              'Start with 2
          Do
            vCount += 1
            vName = vParam.Name & vCount
          Loop While mvSelectItems.Exists(vName)
        Else
          vName = vParam.Name
        End If
        mvSelectItems.Add(vName, vParam.DataType, vParam.Value, pFWO)
      Next
    End Sub

    Private Function GetContactMailingDocumentsSelectionSQL() As SQLStatement
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("MailingTemplate") Then vWhereFields.Add("cmd.mailing_template", mvParameters("MailingTemplate").Value)
      If mvParameters.Exists("MailingDocumentNumber") Then vWhereFields.Add("mailing_document_number", CDBField.FieldTypes.cftLong, mvParameters("MailingDocumentNumber").Value)
      If mvParameters.Exists("CreatedBy") Then vWhereFields.Add("created_by", mvParameters("CreatedBy").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)
      If mvParameters.Exists("Mailing") Then vWhereFields.Add("cmd.mailing", mvParameters("Mailing").Value)

      'Check Created on dates
      If mvSelectItems.Exists("date") Then vWhereFields.Add("created_on", CDBField.FieldTypes.cftDate, mvSelectItems("date").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      If mvSelectItems.Exists("date2") Then vWhereFields.Add("created_on#2", CDBField.FieldTypes.cftDate, mvSelectItems("date2").Value, CDBField.FieldWhereOperators.fwoLessThanEqual)

      Dim vAnsiJoins As New AnsiJoins

      If mvParameters.ParameterExists("Fulfilled").Bool = False Then
        vWhereFields.Add("fulfillment_number", "", CDBField.FieldWhereOperators.fwoNullOrEqual)     'IS NULL
      ElseIf mvParameters.ParameterExists("Fulfilled").Bool = True Then
        vWhereFields.Add("cmd.fulfillment_number", "", CDBField.FieldWhereOperators.fwoNotEqual)     'IS NOT NULL
        If mvParameters.Exists("FulfillmentNumber") Then vWhereFields.Add("cmd.fulfillment_number#2", mvParameters("FulfillmentNumber").LongValue)
        If mvParameters.Exists("FulfilledBy") Then vWhereFields.Add("fulfilled_by", mvParameters("FulfilledBy").Value)
        vAnsiJoins.Add("fulfillment_history fh", "cmd.fulfillment_number", "fh.fulfillment_number")
        If mvSelectItems.Exists("date3") Then
          vWhereFields.Add("fulfilled_on", CDBField.FieldTypes.cftDate, mvSelectItems("date3").Value, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If
        If mvSelectItems.Exists("date4") Then
          'Since fulfillment_history.fulfilled_on will contain time data we have to do some trickery and add one day to the end of the date range.
          'We then retrieve those records where fulfilled_on < that calculated date.
          'We do this rather than finding those where fulfilled_on <= the entered end date because that doesn't work due to the attribute containing time data.
          vWhereFields.Add("fulfilled_on#2", CDBField.FieldTypes.cftDate, DateAdd("d", 1, mvSelectItems("date4").Value).ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoLessThan)
        End If
      End If
      vAnsiJoins.Add("mailing_templates mt", "cmd.mailing_template", "mt.mailing_template")
      vAnsiJoins.Add("contacts c", "cmd.contact_number", "c.contact_number")
      vAnsiJoins.Add("mailings m", "cmd.mailing", "m.mailing")

      Return (New SQLStatement(mvEnv.Connection, GetSelectedAttributes, "contact_mailing_documents cmd", vWhereFields, "cmd.mailing_document_number", vAnsiJoins))
    End Function

    Private Function GetMailingSelectionSQL() As SQLStatement
      Dim vWhereFields As New CDBFields
      If mvParameters.Exists("Mailing") Then vWhereFields.Add("m.mailing", mvParameters("Mailing").Value)
      If mvParameters.Exists("MailingDate") Then vWhereFields.Add("mailing_date", CDBField.FieldTypes.cftDate, mvParameters("MailingDate").Value)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("mailings m", "mh.mailing", "m.mailing")
      vAnsiJoins.Add("email_jobs ej", "mh.mailing_number", "ej.mailing_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      Dim vAttrs As String = GetSelectedAttributes.Replace("mailing_number", "mh.mailing_number AS mailing_number")
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMailingHistoryTopic) Then
        If mvParameters.Exists("Topic") Then vWhereFields.Add("mh.topic", mvParameters("Topic").Value)
        If mvParameters.Exists("SubTopic") Then vWhereFields.Add("mh.sub_topic", mvParameters("SubTopic").Value)
        If mvParameters.Exists("Subject") Then vWhereFields.Add("mh.subject", mvParameters("Subject").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)

        vAnsiJoins.AddLeftOuterJoin("topics t", "mh.topic", "t.topic")
        vAnsiJoins.AddLeftOuterJoin("sub_topics st", "mh.topic", "st.topic", "mh.sub_topic", "st.sub_topic")
      Else
        vAttrs = vAttrs.Replace("mh.topic,topic_desc,mh.sub_topic,sub_topic_desc,mh.subject,", "")
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection, vAttrs, "mailing_history mh", vWhereFields, "mailing_number desc", vAnsiJoins)

      If Not mvParameters.HasValue("Mailing") Then
        vAttrs = "null,null,ej.amended_on as mailing_date,ej.mailing_number AS mailing_number,ej.amended_by as mailing_by,null,number_of_emails,number_processed,number_failed,null,null,null,null,mh.topic,topic_desc,mh.sub_topic,sub_topic_desc,mh.subject,email_job_number"
        Dim vWhereFields2 As New CDBFields(New CDBField("mh.mailing"))
        If mvParameters.Exists("MailingDate") Then vWhereFields2.Add("ej.amended_on", CDBField.FieldTypes.cftDate, mvParameters("MailingDate").Value)
        Dim vAnsiJoins2 As New AnsiJoins()
        vAnsiJoins2.Add("mailing_history mh", "ej.mailing_number", "mh.mailing_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbMailingHistoryTopic) Then
          If mvParameters.Exists("Topic") Then vWhereFields2.Add("mh.topic", mvParameters("Topic").Value)
          If mvParameters.Exists("SubTopic") Then vWhereFields2.Add("mh.sub_topic", mvParameters("SubTopic").Value)
          If mvParameters.Exists("Subject") Then vWhereFields2.Add("mh.subject", mvParameters("Subject").Value, CDBField.FieldWhereOperators.fwoLikeOrEqual)

          vAnsiJoins2.AddLeftOuterJoin("topics t", "mh.topic", "t.topic")
          vAnsiJoins2.AddLeftOuterJoin("sub_topics st", "mh.topic", "st.topic", "mh.sub_topic", "st.sub_topic")
        Else
          vAttrs = vAttrs.Replace("mh.topic,topic_desc,mh.sub_topic,sub_topic_desc,mh.subject,", "")
        End If
        Dim vSQL2 As New SQLStatement(mvEnv.Connection, vAttrs, "email_jobs ej", vWhereFields2, "", vAnsiJoins2)
        vSQL.AddUnion(vSQL2)
      End If
      Return vSQL
    End Function

    Private Function GetEventPersonnelAppointmentSelectionSQL() As SQLStatement
      Dim vWhereFields As New CDBFields()
      Dim vStartDate As Date = Date.Parse(mvParameters("Date").Value)
      Dim vEndDate As Date = Date.Parse(mvParameters("Date2").Value)
      vWhereFields.Add("start_date", CDBField.FieldTypes.cftDate, vEndDate.AddDays(1).ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoLessThanEqual)
      vWhereFields.Add("end_date", CDBField.FieldTypes.cftDate, vStartDate.ToString(CAREDateFormat), CDBField.FieldWhereOperators.fwoGreaterThanEqual)
      Dim vAnsiJoins As New AnsiJoins()
      Dim vPersonnelSQL As SQLStatement = GetEventPersonnelSelectionSQL()
      vPersonnelSQL.OrderBy = ""      'Remove order by from the sub query
      vAnsiJoins.Add(String.Format("({0}) ep", vPersonnelSQL.SQL), "ca.contact_number", "ep.contact_number")
      Return New SQLStatement(mvEnv.Connection, "ca.contact_number,start_date,end_date", "contact_appointments ca", vWhereFields, "ca.contact_number", vAnsiJoins)
    End Function

    Private Function GetEventPersonnelSelectionSQL() As SQLStatement
      Dim vWhereFields As New CDBFields()
      AddWhereFieldFromParameter(vWhereFields, "Surname", "surname")
      AddWhereFieldFromParameter(vWhereFields, "Initials", "initials")
      If mvParameters.HasValue("Activity1") Then
        Dim vSubWhereFields As New CDBFields
        Dim vIndex As Integer = 1
        Do
          Dim vFWO As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoOpenBracket
          If vIndex > 1 Then vFWO = CDBField.FieldWhereOperators.fwoOpenBracket Or CDBField.FieldWhereOperators.fwoOR

          vSubWhereFields.Add("activity#" & vIndex, mvParameters("Activity" & vIndex).Value, vFWO)
          vSubWhereFields.Add("activity_value#" & vIndex, mvParameters("ActivityValue" & vIndex).Value, CDBField.FieldWhereOperators.fwoCloseBracket)
          vIndex += 1
        Loop While mvParameters.HasValue("Activity" & vIndex)
        Dim vSubSQL As New SQLStatement(mvEnv.Connection, "contact_number", "contact_categories", vSubWhereFields)
        vWhereFields.Add("p.contact_number", vSubSQL.SQL, CDBField.FieldWhereOperators.fwoIn)
      End If
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "c.contact_number", "p.contact_number")
      Return New SQLStatement(mvEnv.Connection, "c.contact_number,address_number,surname,initials,surname_prefix", "personnel p", vWhereFields, "surname,initials", vAnsiJoins)
    End Function

    Private Function GetExamPersonnelSelectionSQL() As SQLStatement
      Dim vWhereFields As New CDBFields()
      AddWhereFieldFromParameter(vWhereFields, "Surname", "surname")
      AddWhereFieldFromParameter(vWhereFields, "Initials", "initials")
      AddWhereFieldFromParameter(vWhereFields, "Forenames", "forenames")
      AddWhereFieldFromParameter(vWhereFields, "ContactNumber", "ep.contact_number")
      AddWhereFieldFromParameter(vWhereFields, "ExamPersonnelType", "ep.exam_personnel_type")
      vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
      vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("contacts c", "c.contact_number", "ep.contact_number")
      vAnsiJoins.Add("addresses a", "a.address_number", "ep.address_number")
      vAnsiJoins.Add("exam_personnel_types ept", "ept.exam_personnel_type", "ep.exam_personnel_type")
      If mvParameters.Exists("ExamAssessmentType") Then
        Dim vStringList As New StringList(mvParameters("ExamAssessmentType").Value)
        vWhereFields.Add("exam_assessment_type", CDBField.FieldTypes.cftCharacter, vStringList.InList, CDBField.FieldWhereOperators.fwoIn)
        vAnsiJoins.Add("exam_personnel_assess_types epat", "epat.exam_personnel_id", "ep.exam_personnel_id")
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection, "ep.exam_personnel_id,c.contact_number,ep.address_number,surname,forenames,initials,ep.valid_from,ep.valid_to,ep.exam_personnel_type,exam_personnel_type_desc,surname_prefix", "exam_personnel ep", vWhereFields, "surname,initials", vAnsiJoins)
      vSQL.Distinct = True
      Return vSQL
    End Function

#Region "Text Search"

    Private Function GetTextSearchSelectionSQL() As SQLStatement
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add("search_data sd", "ft.[key]", "sd.id")
      vAnsiJoins.Add("contacts c", "sd.contact_number", "c.contact_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("organisations o", "sd.contact_number", "o.organisation_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      vAnsiJoins.Add("events e", "sd.event_number", "e.event_number", AnsiJoin.AnsiJoinTypes.LeftOuterJoin)
      Dim vTable As String = String.Format("containstable(search_data,*,'{0}', 50 ) ft", mvParameters("SearchText").Value)
      Return New SQLStatement(mvEnv.Connection, "rank as rank_number,sd.contact_number,sd.event_number,sd.document_number,isnull(isnull(sd.contact_number,sd.event_number),sd.document_number) AS id_number,isnull( isnull( label_name, name),event_desc) AS id_desc,description AS full_text", vTable, Nothing, "rank desc", vAnsiJoins)
    End Function

    Private Sub ProcessTextSearch(ByVal pDataTable As CDBDataTable)

      Dim vTableList As New List(Of TextSearchItem)
      Dim vTextSearchItem As New TextSearchItem(mvEnv)
      Dim vTable As DataTable = vTextSearchItem.GetDataTable()
      If vTable.Rows.Count = 0 Then RaiseError(DataAccessErrors.daeTextSearchItemsMissing)
      For Each vRow As DataRow In vTable.Rows
        Dim vNewItem As New TextSearchItem(mvEnv)
        vNewItem.InitFromDataRow(vRow)
        vTableList.Add(vNewItem)
      Next

      Dim vMaxRecords As Integer = IntegerValue(mvEnv.GetConfig("sc_text_search_max_select", "500"))
      Dim vSearchTerm As String = mvParameters("SearchText").Value
      Dim vSearchType As String = mvParameters("SearchType").Value
      If Not vSearchType.StartsWith("A") Then     'If not all then restrict
        For vIndex As Integer = vTableList.Count - 1 To 0 Step -1
          If Not vTableList(vIndex).ItemType.StartsWith(vSearchType.Substring(0, 1), StringComparison.CurrentCultureIgnoreCase) Then
            vTableList.Remove(vTableList(vIndex))
          End If
        Next
      End If
      Dim vItems As New ExpressionEvaluator(vSearchTerm)
      If vItems.Invalid Then RaiseError(DataAccessErrors.daeInvalidTextSearch)
      'Raise an error if there are not table(s) in the vTableList
      If vTableList.Count = 0 Then RaiseError(DataAccessErrors.daeTextSearchItemsMissing)
      vSearchTerm = vItems.Expression
      Debug.Print("Search for:" & vSearchTerm)
      'RankNumber,ItemNumber,Description,ItemType,ItemSource,ItemText
      Dim vSQL As SQLStatement = GetAllSearchTableSQL(vTableList, vSearchTerm, vMaxRecords)
      Dim vFirstTable As DataTable = vSQL.GetDataTable
      ProcessItemText(vFirstTable)                            'Remove returns and unwanted chars
      vItems.MergeTableRows(vFirstTable)                      'Merge records for the same target record

      If vFirstTable.Rows.Count < vMaxRecords AndAlso vItems.ParseExpression > 1 Then
        mvTables = vTableList
        mvDataSet = vFirstTable.DataSet
        Dim vDataTable As DataTable = vItems.EvaluateExpression(AddressOf GetExpressionDataTable)
        If vDataTable IsNot Nothing Then vFirstTable = vItems.JoinTables(vFirstTable, vDataTable, ExpressionEvaluator.BooleanQueryJoinType.OrJoin)
        vItems.MergeTableRows(vFirstTable)                      'Merge records for the same target record
      End If
      pDataTable.SetEnvironment(mvEnv)
      Dim vCount As Integer
      For Each vRow As DataRow In vFirstTable.Rows
        pDataTable.AddRowFromItems(vRow.ItemArray)
        vCount += 1
        If vCount >= vMaxRecords Then Exit For
      Next
      pDataTable.SuppressData()
    End Sub

    Private mvTables As List(Of TextSearchItem)
    Private mvDataSet As DataSet

    Private Function GetExpressionDataTable(ByVal pSearchTerm As String, ByVal pName As String) As DataTable
      Dim vTable As DataTable = GetAllSearchTableSQL(mvTables, pSearchTerm, 0).GetDataTable
      vTable.DataSet.Tables.Remove(vTable)
      vTable.TableName = pName
      ProcessItemText(vTable)
      mvDataSet.Tables.Add(vTable)
      Return vTable
    End Function

    Private Sub ProcessItemText(ByVal pTable As DataTable)
      For Each vRow As DataRow In pTable.Rows
        vRow("item_text") = vRow("item_text").ToString.Replace(Chr(11), "")    'Fix garbage data from RHS database
        vRow("item_text") = vRow("item_text").ToString.Replace(vbLf, "~")      'Remove returns
      Next
    End Sub

    Private Function GetAllSearchTableSQL(ByVal pSearchItems As List(Of TextSearchItem), ByVal pSearchTerm As String, ByVal pMaxRecords As Integer) As SQLStatement
      Dim vSQL As New SQLStatement(mvEnv.Connection, "*", "", New CDBFields, "rank desc")
      vSQL.MaxRows = pMaxRecords
      For Each vSearchTable As TextSearchItem In pSearchItems
        vSQL.AddUnion(GetSearchTableSQL(vSearchTable, pSearchTerm))
      Next
      Return vSQL
    End Function

    Private Function GetSearchTableSQL(ByVal pSearchItem As TextSearchItem, ByVal pSearchTerm As String) As SQLStatement
      Dim vTable As String = String.Format("containstable({0},*,'{1}') st", pSearchItem.TableName, pSearchTerm)
      Dim vAnsiJoins As New AnsiJoins()
      vAnsiJoins.Add(pSearchItem.TableName & " t1", "st.[key]", "t1." & pSearchItem.KeyAttributeName)
      If pSearchItem.LinkTableName.Length > 0 Then
        If pSearchItem.TableName = "addresses" Then
          vAnsiJoins.Add("contact_addresses lt2", "t1.address_number", "lt2.address_number")
          vAnsiJoins.Add(pSearchItem.LinkTableName & " lt", "lt2.contact_number", "lt." & pSearchItem.LinkAttributeName)
        Else
          vAnsiJoins.Add(pSearchItem.LinkTableName & " lt", "t1." & pSearchItem.UniqueAttributeName, "lt." & pSearchItem.LinkAttributeName)
        End If
      End If
      Dim vFields As New StringBuilder
      With vFields
        .Append("rank, ")
        If pSearchItem.TableName = "addresses" Then
          .Append("lt.")
        Else
          .Append("t1.")
        End If
        .Append(pSearchItem.UniqueAttributeName)
        .Append(" AS item_number, ")
        If pSearchItem.ItemType = "Contact" Then
          .Append("ownership_group, ")
        Else
          .Append("department AS ownership_group, ")
        End If
        .Append(pSearchItem.DescriptionAttributeName)
        .Append(" AS description, ")
        .Append("'")
        .Append(pSearchItem.ItemType)
        .Append("' AS item_type,  ")
        .Append("'")
        .Append(pSearchItem.ItemSource)
        .Append("' AS item_source, ")
        .Append(pSearchItem.TextAttrsAsField)
        .Append(" AS item_text")
      End With
      Return New SQLStatement(mvEnv.Connection, vFields.ToString, vTable, Nothing, "", vAnsiJoins)
    End Function

#End Region

  End Class
End Namespace

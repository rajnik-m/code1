Namespace Access
  Partial Public Class DataSelection

    Private Enum QBETypes As Integer
      Contacts
      Organisations
      Events
    End Enum

    Private Enum QBERowTypes As Integer
      Include
      Exclude
    End Enum

    Private mvViewName As String

    Private Sub InitQueryByExampleContacts()
      Dim vColumns As New StringBuilder

      mvViewName = GetViewName("QC")
      If mvViewName.Length > 0 Then
        Dim vTable As DataTable = mvEnv.Connection.GetAttributeNames(mvViewName)
        Dim vFirst As Boolean = True
        For Each vRow As DataRow In vTable.Rows
          Dim vName As String = ProperName(vRow("COLUMN_NAME").ToString)
          If Not vFirst Then vColumns.Append(",")

          Select Case vName
            Case "Country"
              vName = "CountryCode"
            Case "DobEstimated"
              vName = "DOBEstimated"
            Case "ContactGroup"
              vName = "GroupCode"
            Case "NiNumber"
              vName = "NINumber"
            Case "ContactVatCategory"
              vName = "VATCategory"
          End Select
          vColumns.Append(vName)
          vFirst = False
        Next
        vColumns.Append(",AddressLine,ContactName")
      Else
        vColumns.Append("ContactNumber,AddressNumber,ContactName,Title,Initials,Forenames,Surname,Honorifics,Salutation,LabelName")
        vColumns.Append(",PreferredForename,Sex,DateOfBirth,Source,SourceDate,Status,StatusDate,StatusReason,StatusDesc")
        vColumns.Append(",HouseName,Address,AddressLine1,AddressLine2,AddressLine3,AddressLine4,Town,County,Postcode,CountryCode,Branch,AddressLine,ContactType")
        vColumns.Append(",Notes,AmendedOn,AmendedBy,OwnershipGroup,DOBEstimated,GroupCode,Department,VATCategory,NINumber")
        vColumns.Append(",PrefixHonorifics,SurnamePrefix,InformalSalutation,BuildingNumber,AddressType")
        vColumns.Append(",Position,Name")
      End If
      mvResultColumns = vColumns.ToString
      mvSelectColumns = "ContactNumber,Title,Forenames,Surname,Position,Name,Town,Postcode"
      mvHeadings = "Contact Number,Title,Forenames,Surname,Position,Name,Town,Postcode"
      mvWidths = "1000,1000,1000,1000,1000,1000,1000,1000"
      mvRequiredItems = "ContactNumber,OwnershipGroup"
      mvDescription = "Contact QBE"
      mvCode = "QBEC"
    End Sub

    Private Sub InitQueryByExampleOrganisations()
      Dim vColumns As New StringBuilder

      mvViewName = GetViewName("QO")
      If mvViewName.Length > 0 Then
        Dim vTable As DataTable = mvEnv.Connection.GetAttributeNames(mvViewName)
        Dim vFirst As Boolean = True
        For Each vRow As DataRow In vTable.Rows
          Dim vName As String = ProperName(vRow("COLUMN_NAME").ToString)
          If Not vFirst Then vColumns.Append(",")

          Select Case vName
            Case "Country"
              vName = "CountryCode"
            Case "OrganisationGroup"
              vName = "GroupCode"
          End Select
          vColumns.Append(vName)
          vFirst = False
        Next
        vColumns.Append(",AddressLine")
      Else
        vColumns.Append("OrganisationNumber,AddressNumber,Name,Salutation,LabelName")
        vColumns.Append(",Source,SourceDate,Status,StatusDate,StatusReason,StatusDesc")
        vColumns.Append(",HouseName,Address,AddressLine1,AddressLine2,AddressLine3,AddressLine4,Town,County,Postcode,CountryCode,Branch,AddressLine")
        vColumns.Append(",Notes,AmendedOn,AmendedBy,OwnershipGroup,GroupCode,Department")
        vColumns.Append(",BuildingNumber,AddressType")
      End If
      mvResultColumns = vColumns.ToString
      mvSelectColumns = "OrganisationNumber,Name,Town,Postcode"
      mvHeadings = "Organisation Number,Name,Town,Postcode"
      mvWidths = "1000,1000,1000,1000"
      mvRequiredItems = "OrganisationNumber,OwnershipGroup"
      mvDescription = "Organisation QBE"
      mvCode = "QBEO"
    End Sub

    Private Sub InitQueryByExampleEvents()
      Dim vColumns As New StringBuilder

      mvViewName = GetViewName("QE")
      If mvViewName.Length > 0 Then
        Dim vTable As DataTable = mvEnv.Connection.GetAttributeNames(mvViewName)
        Dim vFirst As Boolean = True
        For Each vRow As DataRow In vTable.Rows
          Dim vName As String = ProperName(vRow("COLUMN_NAME").ToString)
          If Not vFirst Then vColumns.Append(",")
          vColumns.Append(vName)
          vFirst = False
        Next
      Else
        vColumns.Append("EventNumber,EventDesc,EventReference,LongDescription,EventClass,EventStatus,Subject,SkillLevel,StartDate,StartTime,EndDate,EndTime")
        vColumns.Append(",Branch,Source,Venue,VenueReference,Location,Notes")
        vColumns.Append(",NumberOfAttendees,MinimumAttendees,MaximumAttendees,TargetAttendees,NumberOnWaitingList,MaximumOnWaitingList,NumberInterested,NumberOfBookings")
        vColumns.Append(",FreeOfCharge,MultiSession,Booking,BookingsClose,NameAttendees,EventPricingMatrix,WaitingListControlMethod,ChargeForWaiting,BalanceBookingFee")
        vColumns.Append(",BalanceBookingDue,MinimumSponsorshipAmount,SponsorshipDue,PledgedAmountDue")
        vColumns.Append(",Template,MoveSessionDates,ActivityGroup,RelationshipGroup,Department,SponsorshipProduct,SponsorshipRate")
        vColumns.Append(",TargetIncome,SponsoredCosts,SponsorshipIncome,DonationIncome,BookingIncome,OtherIncome,TotalIncome,TotalCosts,TotalExpenditure,ReturnOnInvestment")
        vColumns.Append(",DelegateContribution,GiftAidDeclarationCount,GiftAidDeclarationValue,FinancialLastCalculated,SponsorshipLastCalculated")
      End If
      mvResultColumns = vColumns.ToString
      mvSelectColumns = "EventNumber,EventDesc,EventReference,StartDate,EndDate,Subject,SkillLevel,EventStatus"
      mvHeadings = "Event Number,Description,Reference,Start Date,End Date,Subject,Skill Level,Event Status"
      mvWidths = "1000,1000,1000,1000,1000,1000,1000,1000"
      mvRequiredItems = "EventNumber"
      mvDescription = "Event QBE"
      mvCode = "QBEE"
    End Sub

    Private Function GetViewName(pViewType As String) As String
      Dim vWhereFields As New CDBFields
      vWhereFields.AddClientDeptLogname(mvEnv.ClientCode, mvEnv.User.Department, mvEnv.User.Logname)
      vWhereFields.Add("view_type", pViewType)
      Dim vSQL As New SQLStatement(mvEnv.Connection, "view_name", "view_names", vWhereFields)
      vSQL.SetOrderByClientDeptLogname("view_name")
      Return vSQL.GetValue
    End Function

    Private Sub GetQueryByExampleContacts(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()

      Dim vTempTable As String = "smcam_smapp_" & mvParameters("SelectionSetNumber").Value

      If Not mvParameters.ContainsKey("ViewSelectedRecords") Then
        vWhereFields.Add("contact_group", mvParameters("ContactGroup").Value)
        vWhereFields.Add("contact_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
        Dim vExcludeStatements As New List(Of SQLStatement)
        ProcessPageParameters("QECO", vWhereFields, vAnsiJoins, vExcludeStatements)     'Contact
        ProcessPageParameters("QECA", vWhereFields, vAnsiJoins, vExcludeStatements)     'Addresses
        ProcessPageParameters("QECN", vWhereFields, vAnsiJoins, vExcludeStatements)     'Numbers
        ReplaceCurrentParameter("QECC")
        ProcessPageParameters("QECC", vWhereFields, vAnsiJoins, vExcludeStatements)     'Categories
        ProcessPageParameters("QECM", vWhereFields, vAnsiJoins, vExcludeStatements)     'Members
        ProcessPageParameters("QECP", vWhereFields, vAnsiJoins, vExcludeStatements)     'Payment Plans
        ProcessPageParameters("QECJ", vWhereFields, vAnsiJoins, vExcludeStatements)     'Positions
        ProcessPageParameters("QECR", vWhereFields, vAnsiJoins, vExcludeStatements)     'Roles
        ReplaceCurrentParameter("QECS")
        ProcessPageParameters("QECS", vWhereFields, vAnsiJoins, vExcludeStatements)     'Suppressions
        ProcessPageParameters("QECL", vWhereFields, vAnsiJoins, vExcludeStatements)     'Mailings
        ProcessPageParameters("QECD", vWhereFields, vAnsiJoins, vExcludeStatements)     'Documents
        ProcessPageParameters("QECF", vWhereFields, vAnsiJoins, vExcludeStatements)     'Financial
        ProcessPageParameters("QECE", vWhereFields, vAnsiJoins, vExcludeStatements)     'Events
        ProcessPageParameters("QECX", vWhereFields, vAnsiJoins, vExcludeStatements)     'Exams
        ProcessPageParameters("QECH", vWhereFields, vAnsiJoins, vExcludeStatements)     'RelationshipsFrom
        ProcessPageParameters("QECT", vWhereFields, vAnsiJoins, vExcludeStatements)     'RelationshipsTo
        ProcessPageParameters("QECU", vWhereFields, vAnsiJoins, vExcludeStatements)     'Fundraising Requests
        Dim vIndex As Integer
        For Each vExclude As SQLStatement In vExcludeStatements
          vWhereFields.Add("Exclude" & vIndex, vExclude.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
          vIndex += 1
        Next
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, String.Format("{0},1,c.contact_number,c.address_number", mvParameters("SelectionSetNumber").Value), "contacts c", vWhereFields, "c.contact_number", vAnsiJoins)
        vSQLStatement.Distinct = True
        vSQLStatement.UseAnsiSQL = True
        Debug.Print(vSQLStatement.SQL)

        Dim vSQL As String = String.Format("INSERT INTO {0}( selection_set, revision, contact_number, address_number ){1}", vTempTable, vSQLStatement.SQL)
        mvEnv.Connection.DeleteAllRecords(vTempTable)
        mvEnv.Connection.ExecuteSQL(vSQL, CDBConnection.cdbExecuteConstants.sqlShowError)
      End If

      Dim vList As New StringList(mvSelectColumns, StringSplitOptions.RemoveEmptyEntries)
      vAnsiJoins.Clear()
      vAnsiJoins.Add("contacts c", "sc.contact_number", "c.contact_number")
      vAnsiJoins.Add("addresses a", "c.address_number", "a.address_number")
      If vList.Contains("AddressLine") Then
        vAnsiJoins.Add("countries co", "co.country", "a.country")
      End If
      vAnsiJoins.AddLeftOuterJoin("contact_positions cp", "c.contact_number", "cp.contact_number", "c.address_number", "cp.address_number")
      vAnsiJoins.AddLeftOuterJoin("organisations o", "cp.organisation_number", "o.organisation_number")

      Dim vAttrs As New StringBuilder
      Dim vAddComma As Boolean
      For Each vItem As String In vList
        If vAddComma Then vAttrs.Append(",")
        Select Case vItem
          Case "ContactNumber", "AddressNumber", "Department", "AmendedBy", "AmendedOn", "Notes", "OwnershipGroup", "Source", "SourceDate", "Status", "StatusDate", "StatusReason"
            vAttrs.Append("c.")
            vAttrs.Append(AttributeName(vItem))
          Case "AddressLine1", "AddressLine2", "AddressLine3", "AddressLine4"
            vAttrs.Append("address_line")
            vAttrs.Append(vItem(vItem.Length - 1))
          Case "BranchCode"
            vAttrs.Append("branch")
          Case "CountryCode"
            vAttrs.Append("a.country")
          Case "DOBEstimated"
            vAttrs.Append("dob_estimated")
          Case "GroupCode"
            vAttrs.Append("c.contact_group")
          Case "NINumber"
            vAttrs.Append("ni_number")
          Case "VATCategory"
            vAttrs.Append("contact_vat_category")
          Case "ContactName"
            vAttrs.Append("CONTACT_NAME")
          Case "AddressLine"
            vAttrs.Append("ADDRESS_LINE")
          Case "RgbStatus"
            vAttrs.Append("s.rgb_value AS RgbStatus")
          Case Else
            vAttrs.Append(AttributeName(vItem))
        End Select
        vAddComma = True
      Next
      Dim vItems As String = vAttrs.ToString

      If vList.Contains("Status") OrElse vList.Contains("StatusDesc") Then
        vAnsiJoins.AddLeftOuterJoin("statuses s", "c.status", "s.status")
        If Not vList.Contains("RgbStatus") Then
          vAttrs.Append(",s.rgb_value AS RgbStatus")
          pDataTable.AddColumn("RgbStatus", CDBField.FieldTypes.cftInteger)
          vItems = vItems & ",RgbStatus"
          mvSelectColumns &= ",RgbStatus"
          mvHeadings &= ","
          mvWidths &= ",1"
        End If
      End If
      If vList.Contains("ContactName") Then
        Dim vContact As New Contact(mvEnv)
        AddFieldsFromList(vAttrs, vContact.GetRecordSetFieldsName)
      End If
      If vList.Contains("AddressLine") Then
        Dim vAddress As New Address(mvEnv)
        AddFieldsFromList(vAttrs, vAddress.GetRecordSetFieldsCountry)

      End If
      Dim vFields As String = vAttrs.ToString
      vFields = vFields.Replace("CONTACT_NAME", "")
      vFields = vFields.Replace("ADDRESS_LINE", "")
      If mvViewName.Length > 0 Then
        vAnsiJoins.Clear()
        vAnsiJoins.Add(mvViewName & " c", "sc.contact_number", "c.contact_number")
        vFields = vFields.Replace("s.rgb_value", "rgb_value")
        vFields = vFields.Replace("a.country", "country")
        vFields = vFields.Replace("a.address_number", "c.address_number")
        vFields = vFields.Replace("a.branch", "branch")
      End If
      Dim vReturnSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields.TrimStart(","c)), String.Format("{0} sc", vTempTable), New CDBFields, "surname,c.contact_number", vAnsiJoins)
      vItems = vItems.Replace("c.contact_number", "DISTINCT_CONTACT_NUMBER")
      pDataTable.FillFromSQL(mvEnv, vReturnSQL, vItems)
      pDataTable.SuppressData()
    End Sub

    Private Sub GetQueryByExampleOrganisations(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("organisation_group", mvParameters("OrganisationGroup").Value.ToString)
      Dim vTempTable As String = "smcam_smapp_" & mvParameters("SelectionSetNumber").Value
      Dim vExcludeStatements As New List(Of SQLStatement)
      If Not mvParameters.ContainsKey("ViewSelectedRecords") Then
        ProcessPageParameters("QEOO", vWhereFields, vAnsiJoins, vExcludeStatements)     'Contact
        ProcessPageParameters("QEOA", vWhereFields, vAnsiJoins, vExcludeStatements)     'Addresses
        ProcessPageParameters("QEON", vWhereFields, vAnsiJoins, vExcludeStatements)     'Numbers
        ReplaceCurrentParameter("QEOC")
        ProcessPageParameters("QEOC", vWhereFields, vAnsiJoins, vExcludeStatements)     'Categories
        ProcessPageParameters("QEOM", vWhereFields, vAnsiJoins, vExcludeStatements)     'Members
        ProcessPageParameters("QEOP", vWhereFields, vAnsiJoins, vExcludeStatements)     'Payment Plans
        ProcessPageParameters("QEOJ", vWhereFields, vAnsiJoins, vExcludeStatements)     'Positions
        ProcessPageParameters("QEOR", vWhereFields, vAnsiJoins, vExcludeStatements)     'Roles
        ReplaceCurrentParameter("QEOS")
        ProcessPageParameters("QEOS", vWhereFields, vAnsiJoins, vExcludeStatements)     'Suppressions
        ProcessPageParameters("QEOL", vWhereFields, vAnsiJoins, vExcludeStatements)     'Mailings
        ProcessPageParameters("QEOD", vWhereFields, vAnsiJoins, vExcludeStatements)     'Documents
        ProcessPageParameters("QEOF", vWhereFields, vAnsiJoins, vExcludeStatements)     'Financial
        ProcessPageParameters("QEOE", vWhereFields, vAnsiJoins, vExcludeStatements)     'Events
        ProcessPageParameters("QEOH", vWhereFields, vAnsiJoins, vExcludeStatements)     'RelationshipsFrom
        ProcessPageParameters("QEOT", vWhereFields, vAnsiJoins, vExcludeStatements)     'RelationshipsTo
        ProcessPageParameters("QEOU", vWhereFields, vAnsiJoins, vExcludeStatements)     'Fundraising Requests

        Dim vIndex As Integer
        For Each vExclude As SQLStatement In vExcludeStatements
          vWhereFields.Add("Exclude" & vIndex, vExclude.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
          vIndex += 1
        Next
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, String.Format("{0},1,o.organisation_number,o.address_number", mvParameters("SelectionSetNumber").Value), "organisations o", vWhereFields, "o.organisation_number", vAnsiJoins)
        vSQLStatement.Distinct = True
        vSQLStatement.UseAnsiSQL = True
        Debug.Print(vSQLStatement.SQL)
        Dim vSQL As String = String.Format("INSERT INTO {0}( selection_set, revision, contact_number, address_number ){1}", vTempTable, vSQLStatement.SQL)
        mvEnv.Connection.DeleteAllRecords(vTempTable)
        mvEnv.Connection.ExecuteSQL(vSQL, CDBConnection.cdbExecuteConstants.sqlShowError)
      End If

      Dim vList As New StringList(mvSelectColumns, StringSplitOptions.RemoveEmptyEntries)
      vAnsiJoins.Clear()
      vAnsiJoins.Add("organisations o", "sc.contact_number", "o.organisation_number")
      vAnsiJoins.Add("contacts c", "o.contact_number", "c.contact_number")
      vAnsiJoins.Add("addresses a", "o.address_number", "a.address_number")
      If vList.Contains("AddressLine") Then
        vAnsiJoins.Add("countries co", "co.country", "a.country")
      End If

      Dim vAttrs As New StringBuilder
      Dim vAddComma As Boolean
      For Each vItem As String In vList
        If vAddComma Then vAttrs.Append(",")
        Select Case vItem
          Case "OrganisationNumber", "AddressNumber", "Department", "AmendedBy", "AmendedOn", "Notes", "OwnershipGroup", "Source", "SourceDate", "Status", "StatusDate", "StatusReason"
            vAttrs.Append("o.")
            vAttrs.Append(AttributeName(vItem))
          Case "AddressLine1", "AddressLine2", "AddressLine3", "AddressLine4"
            vAttrs.Append("address_line")
            vAttrs.Append(vItem(vItem.Length - 1))
          Case "CountryCode"
            vAttrs.Append("a.country")
          Case "GroupCode"
            vAttrs.Append("organisation_group")
          Case "AddressLine"
            vAttrs.Append("ADDRESS_LINE")
          Case Else
            vAttrs.Append(AttributeName(vItem))
        End Select
        vAddComma = True
      Next
      Dim vItems As String = vAttrs.ToString

      If vList.Contains("Status") OrElse vList.Contains("StatusDesc") Then
        'vAnsiJoins.AddLeftOuterJoin("statuses s", "o.status", "s.status")
        If vList.Contains("RgbStatus") Then
          vAttrs.Replace("rgb_status", "s.rgb_value AS rgb_status")
        Else
          vAttrs.Append(",s.rgb_value AS RgbStatus")
          pDataTable.AddColumn("RgbStatus", CDBField.FieldTypes.cftInteger)
          vItems = vItems & ",RgbStatus"
          mvSelectColumns &= ",RgbStatus"
          mvHeadings &= ","
          mvWidths &= ",1"
        End If
      End If
      If vList.Contains("AddressLine") Then
        Dim vAddress As New Address(mvEnv)
        AddFieldsFromList(vAttrs, vAddress.GetRecordSetFieldsCountry)
      End If
      Dim vFields As String = vAttrs.ToString
      vFields = vFields.Replace("ADDRESS_LINE", "")
      If mvViewName.Length > 0 Then
        vAnsiJoins.Clear()
        vAnsiJoins.Add(mvViewName & " o", "sc.contact_number", "o.organisation_number")
        vFields = vFields.Replace("s.rgb_value", "rgb_value")
        vFields = vFields.Replace("a.country", "country")
        vFields = vFields.Replace("a.address_number", "c.address_number")
        vFields = vFields.Replace("a.branch", "branch")
      End If
      Dim vReturnSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields.TrimStart(","c)), String.Format("{0} sc", vTempTable), New CDBFields, "name,organisation_number", vAnsiJoins)
      vItems = vItems.Replace("o.organisation_number", "DISTINCT_ORGANISATION_NUMBER")
      pDataTable.FillFromSQL(mvEnv, vReturnSQL, vItems)
      pDataTable.SuppressData()
    End Sub

    Private Sub GetQueryByExampleEvents(ByVal pDataTable As CDBDataTable)
      Dim vWhereFields As New CDBFields()
      Dim vAnsiJoins As New AnsiJoins()
      vWhereFields.Add("event_group", mvParameters("EventGroup").Value.ToString)
      Dim vTempTable As String = "smcam_smapp_" & mvParameters("SelectionSetNumber").Value
      If Not mvParameters.ContainsKey("ViewSelectedRecords") Then
        Dim vExcludeStatements As New List(Of SQLStatement)
        ProcessPageParameters("QEEH", vWhereFields, vAnsiJoins, vExcludeStatements)     'Event
        ProcessPageParameters("QEED", vWhereFields, vAnsiJoins, vExcludeStatements)     'Booking Details
        ProcessPageParameters("QEEO", vWhereFields, vAnsiJoins, vExcludeStatements)     'Organiser
        ProcessPageParameters("QEES", vWhereFields, vAnsiJoins, vExcludeStatements)     'Sessions
        ProcessPageParameters("QEEV", vWhereFields, vAnsiJoins, vExcludeStatements)     'Venues
        ProcessPageParameters("QEET", vWhereFields, vAnsiJoins, vExcludeStatements)     'Topics
        Dim vIndex As Integer
        For Each vExclude As SQLStatement In vExcludeStatements
          vWhereFields.Add("Exclude" & vIndex, vExclude.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
          vIndex += 1
        Next
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, String.Format("{0},1,e.event_number,e.event_number", mvParameters("SelectionSetNumber").Value), "events e", vWhereFields, "e.event_number", vAnsiJoins)
        vSQLStatement.Distinct = True
        vSQLStatement.UseAnsiSQL = True
        Debug.Print(vSQLStatement.SQL)
        Dim vSQL As String = String.Format("INSERT INTO {0}( selection_set, revision, contact_number, address_number ){1}", vTempTable, vSQLStatement.SQL)
        mvEnv.Connection.DeleteAllRecords(vTempTable)
        mvEnv.Connection.ExecuteSQL(vSQL, CDBConnection.cdbExecuteConstants.sqlShowError)
      End If

      Dim vList As New StringList(mvSelectColumns, StringSplitOptions.RemoveEmptyEntries)
      vAnsiJoins.Clear()
      vAnsiJoins.Add("events e", "sc.contact_number", "e.event_number")
      vAnsiJoins.Add("sessions s", "e.event_number", "s.event_number")
      Dim vAttrs As New StringBuilder
      Dim vAddComma As Boolean
      For Each vItem As String In vList
        If vAddComma Then vAttrs.Append(",")
        Select Case vItem
          Case "EventNumber", "AmendedBy", "AmendedOn", "StartDate", "LongDescription"
            vAttrs.Append("e.")
            vAttrs.Append(AttributeName(vItem))
          Case "EventStatus"
            If mvViewName.Length = 0 Then vAttrs.Append("es.")
            vAttrs.Append(AttributeName(vItem))
          Case Else
            vAttrs.Append(AttributeName(vItem))
        End Select
        vAddComma = True
      Next
      Dim vItems As String = vAttrs.ToString

      If vList.Contains("EventStatus") Then
        vAnsiJoins.AddLeftOuterJoin("event_statuses es", "e.event_status", "es.event_status")
        If Not vList.Contains("RgbEventStatus") Then
          vAttrs.Append(",es.rgb_value AS RgbEventStatus")
          pDataTable.AddColumn("RgbEventStatus", CDBField.FieldTypes.cftInteger)
          vItems = vItems & ",RgbEventStatus"
          mvSelectColumns &= ",RgbEventStatus"
          mvHeadings &= ","
          mvWidths &= ",1"
        End If
      End If

      Dim vFields As String = vAttrs.ToString
      vWhereFields.Clear()
      If mvViewName.Length > 0 Then
        vAnsiJoins.Clear()
        vAnsiJoins.Add(mvViewName & " e", "sc.contact_number", "e.event_number")
        vFields = vFields.Replace("es.rgb_value", "rgb_value")
      Else
        vWhereFields.Add("session_type", "0")
      End If
      Dim vReturnSQL As New SQLStatement(mvEnv.Connection, RemoveBlankItems(vFields), String.Format("{0} sc", vTempTable), vWhereFields, "e.start_date DESC", vAnsiJoins)
      vItems = vItems.Replace("e.event_number", "DISTINCT_EVENT_NUMBER")
      pDataTable.FillFromSQL(mvEnv, vReturnSQL, vItems)
    End Sub

    Private Sub ReplaceCurrentParameter(ByVal pPageCode As String)
      Dim vFound As Boolean
      Do
        vFound = False
        For Each vParam As CDBParameter In mvParameters
          If vParam.Name.StartsWith(pPageCode) Then
            Dim vLen As Integer = vParam.Name.Substring(5).IndexOf("_")
            Dim vRowNumber As Integer = IntegerValue(vParam.Name.Substring(5, vLen))
            Dim vParamName As String = vParam.Name.Substring(6 + vLen)
            While Char.IsDigit(vParamName(vParamName.Length - 1))
              vParamName = vParamName.Substring(0, vParamName.Length - 1)
            End While
            If vParamName = "Current" Then
              If vParam.Value = "N" Then
                mvParameters.Add(String.Format("{0}_{1}_ValidFrom", pPageCode, vRowNumber), "Adv... >(" & TodaysDate())
                mvParameters.Add(String.Format("{0}_{1}_ValidTo", pPageCode, vRowNumber), "Adv... <OR)" & TodaysDate())
              Else
                mvParameters.Add(String.Format("{0}_{1}_ValidFrom", pPageCode, vRowNumber), "Adv... <=" & TodaysDate())
                mvParameters.Add(String.Format("{0}_{1}_ValidTo", pPageCode, vRowNumber), "Adv... >=" & TodaysDate())
              End If
              mvParameters.Remove(vParam)
              vFound = True
              Exit For
            End If
          End If
        Next
      Loop While vFound
    End Sub

    Private Function AddPageParameters(ByVal pPageCode As String, ByVal pQBERowType As QBERowTypes, ByVal pWhereFields As CDBFields, ByVal pPrefix As String, ByVal pAlternates As CDBParameters, ByRef pRowCount As Integer) As Integer
      Dim vParamName As String
      Dim vRowNumber As Integer
      Dim vCount As Integer
      Dim vDefaultAddressOnly As Boolean = False
      Dim vPrefixRowZeroExists As Boolean = False

      'BR20313 to handle a checkbox that is not an attribute but just controls default address search
      If (pPageCode = "QECA" OrElse pPageCode = "QEOA") Then
        Dim vParamToRemove As String = ""
        For Each vCheckParam As CDBParameter In mvParameters
          If vCheckParam.Name.EndsWith("DefaultAddressOnly") AndAlso vCheckParam.Name.StartsWith(pPageCode) Then
            vParamToRemove = vCheckParam.Name.ToString
            vDefaultAddressOnly = mvParameters(vParamToRemove).Bool
            Exit For
          End If
        Next
        If vParamToRemove <> "" Then
          mvParameters.Remove(vParamToRemove)
        End If
      End If

      For Each vParam As CDBParameter In mvParameters
        Dim vSpecialColumn As Boolean = False
        If vParam.Name.StartsWith(pPageCode) Then
          Dim vLen As Integer = vParam.Name.Substring(5).IndexOf("_")
          vRowNumber = IntegerValue(vParam.Name.Substring(5, vLen))
          vParamName = vParam.Name.Substring(6 + vLen)
          While Char.IsDigit(vParamName(vParamName.Length - 1))
            vParamName = vParamName.Substring(0, vParamName.Length - 1)
          End While
          If vParamName = "Include" OrElse vParamName = "Exclude" Then Continue For
          Select Case pQBERowType
            Case QBERowTypes.Include
              If mvParameters(String.Format("{0}_{1}_Include", pPageCode, vRowNumber)).Value <> "I" Then Continue For
            Case QBERowTypes.Exclude
              If mvParameters(String.Format("{0}_{1}_Include", pPageCode, vRowNumber)).Value <> "E" Then Continue For
          End Select
          Dim vValue As String = vParam.Value
          Dim vOperator As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoEqual
          Dim vFieldType As CDBField.FieldTypes = CDBField.FieldTypes.cftCharacter
          'Temporary solution for data type
          Select Case vParamName
            Case "Accepted", "ActivityDate", "Applied", "BalanceBookingDue", "BalancePaidDate", "BookingsClose", "CancelledOn", "ConfirmedOn", "Dated", "DateOfBirth", _
                 "DepositDate", "DepositPaidDate", "DueDate", "EndDate", "FullPaymentDate", "Joined", "MembershipCardExpires", "NextPaymentDue", "PaymentDate", "PledgedAmountDue", _
                 "SourceDate", "SponsorshipDue", "StartDate", "StatusDate", "TransactionDate", "ValidFrom", "ValidTo", "MailingDate"
              vFieldType = CDBField.FieldTypes.cftDate
            Case "TheirReference", "StatusReason", "Notes", "LongDescription", "Location", "Address", "Precis", "HouseName"
              vFieldType = CDBField.FieldTypes.cftMemo
            Case "Number"
              vSpecialColumn = True
          End Select
          Dim vBetweenValues(1) As String
          If vValue.StartsWith("Adv...") Then
            vValue = vValue.Substring(7).TrimStart
            If vValue.StartsWith("IN") Then
              vOperator = CDBField.FieldWhereOperators.fwoIn
              vValue = vValue.Substring(2).Trim
              vValue = vValue.Substring(1, vValue.Length - 2)
            ElseIf vValue.StartsWith("NOT IN") Then
              vOperator = CDBField.FieldWhereOperators.fwoNotIn
              vValue = vValue.Substring(6).Trim
              vValue = vValue.Substring(1, vValue.Length - 2)
            ElseIf vValue.StartsWith("IS NULL") Then
              vOperator = CDBField.FieldWhereOperators.fwoEqual
              vValue = ""
            ElseIf vValue.StartsWith("IS NOT NULL") Then
              vOperator = CDBField.FieldWhereOperators.fwoNotEqual
              vValue = ""
            ElseIf vValue.StartsWith(">=") Then
              vOperator = CDBField.FieldWhereOperators.fwoGreaterThanEqual
              vValue = vValue.Substring(2).TrimStart
            ElseIf vValue.StartsWith("<=") Then
              vOperator = CDBField.FieldWhereOperators.fwoLessThanEqual
              vValue = vValue.Substring(2).TrimStart
            ElseIf vValue.StartsWith(">") Then
              vOperator = CDBField.FieldWhereOperators.fwoGreaterThan
              vValue = vValue.Substring(1).TrimStart
              If vValue.Substring(0).StartsWith("OR") Then
                vValue = vValue.Substring(3).TrimStart
                vOperator = vOperator Or CDBField.FieldWhereOperators.fwoOR
              End If
              If vValue.Substring(0).StartsWith("(") Then
                vValue = vValue.Substring(1).TrimStart
                vOperator = vOperator Or CDBField.FieldWhereOperators.fwoOpenBracket
              End If
            ElseIf vValue.StartsWith("<") Then
              vOperator = CDBField.FieldWhereOperators.fwoLessThan
              vValue = vValue.Substring(1).TrimStart
              If vValue.Substring(0).StartsWith("OR") Then
                vValue = vValue.Substring(2).TrimStart
                vOperator = vOperator Or CDBField.FieldWhereOperators.fwoOR
              End If
              If vValue.Substring(0).StartsWith(")") Then
                vValue = vValue.Substring(1).TrimStart
                vOperator = vOperator Or CDBField.FieldWhereOperators.fwoCloseBracket
              End If
            ElseIf vValue.StartsWith("BETWEEN") Then
              vOperator = CDBField.FieldWhereOperators.fwoBetweenFrom
              'quick and dirty solution improve later
              vValue = vValue.Substring(7).TrimStart
              vValue = vValue.Replace("AND", "|")
              Dim vWords() As String = vValue.Split("|"c)
              vBetweenValues(0) = vWords(0).Trim
              vBetweenValues(1) = vWords(1).Trim
              vValue = vBetweenValues(0)
            End If
          Else
            If vFieldType = CDBField.FieldTypes.cftCharacter AndAlso vOperator = CDBField.FieldWhereOperators.fwoEqual Then
              vOperator = CDBField.FieldWhereOperators.fwoLikeOrEqual
            ElseIf vFieldType = CDBField.FieldTypes.cftMemo Then
              vOperator = CDBField.FieldWhereOperators.fwoLike
            End If
          End If
          Dim vPrefix As String = pPrefix
          If pAlternates IsNot Nothing Then
            For Each vAlternate As CDBParameter In pAlternates
              Dim vNames As New StringList(vAlternate.Value)
              If vNames.Contains(vParamName) Then vPrefix = vAlternate.Name
            Next
          End If
          Dim vAttrName As String = ""
          Select Case vParamName
            Case "ActivityAndValue"
              vAttrName = "activity" & mvEnv.Connection.DBConcatString & "'-'" & mvEnv.Connection.DBConcatString & "activity_value"
            Case "Balance"
              If vPrefix = "vb" Then vAttrName = "amount"
            Case "BalancePaidDate"
              vAttrName = "payment_date"
            Case "ContactGroup"
              Select Case pPageCode
                Case "QECH", "QECT"
                  Continue For
              End Select
            Case "ContactNumber"
              Select Case pPageCode
                Case "QECH"   'relationship from: This is contact number 1 but the numeric suffix has been removed.  Because of union to organisation links table just called number1
                  vAttrName = "number1"
                Case "QECT"   'relationship to: This is actually contact number 2 but the numeric suffix has been removed
                  vAttrName = "contact_number_2"
              End Select
            Case "Current"
              Select Case vPrefix
                Case "cr", "com"     'Roles  Communications
                  If vParam.Value = "Y" Then
                    vAttrName = "is_active"
                  Else
                    Continue For
                  End If
                Case "ca", "oa", "rt", "rf"
                  If vParam.Value = "Y" Then
                    vAttrName = "historical"
                    vValue = "N"
                  ElseIf vParam.Value = "N" Then
                    vAttrName = "historical"
                    vValue = "Y"
                  Else
                    Continue For
                  End If
                Case "m"
                  vAttrName = "cancellation_reason"
                  If vValue = "Y" Then
                    vOperator = CDBField.FieldWhereOperators.fwoEqual
                    vValue = ""                 'Current members so cancellation reason is NULL
                  Else
                    vOperator = CDBField.FieldWhereOperators.fwoNotEqual
                    vValue = ""
                  End If
                Case Else
                  vSpecialColumn = True
              End Select
            Case "Deposit"
              vAttrName = "deposit_amount"
            Case "DepositPaidDate"
              vAttrName = "deposit_date"
            Case "DOBEstimated"
              vAttrName = "dob_estimated"
            Case "DocumentNumber"
              vAttrName = "communications_log_number"
            Case "DocumentSubject"
              vAttrName = "subject"
            Case "DueDate"
              vAttrName = "full_payment_date"
            Case "Location"
              If vPrefix = "cp" Then vAttrName = "position_location"
            Case "OrganisationGroup"
              Select Case pPageCode
                Case "QEOH", "QEOT"
                  Continue For
              End Select
            Case "OrganisationNumber"
              Select Case pPageCode
                Case "QEOH"   'relationship from: This is organisation number 1 but the numeric suffix has been removed.  Because of union to contact links table just called number1
                  vAttrName = "number1"
                Case "QEOT"   'relationship to: This is actually organisation number 2 but the numeric suffix has been removed
                  vAttrName = "organisation_number_2"
              End Select
            Case "OrganiserReference"
              vAttrName = "reference"
            Case "PaymentPlanNumber"
              vAttrName = "order_number"
            Case "StandardPosition"
              vAttrName = "position"
            Case "STDCode"
              vAttrName = "std_code"
            Case "TotalAmount"
              vAttrName = "full_amount"
            Case "ValidFrom"
              If vPrefix = "cp" Then vAttrName = "started"
            Case "ValidTo"
              If vPrefix = "cp" Then vAttrName = "finished"
            Case "VatCategory"
              vAttrName = "contact_vat_category"
            Case "VatNumber"
              vAttrName = "vat_registration_number"
          End Select
          If String.IsNullOrEmpty(vAttrName) Then vAttrName = AttributeName(vParamName)
          Select Case vPrefix
            Case "c", "o", "e"
              'leave alone
            Case Else
              vPrefix = vPrefix & vRowNumber.ToString
              'BR20313
              If vPrefix.Contains("a0") Then
                vPrefixRowZeroExists = True
              End If
          End Select
          Dim vFieldName As String
          If vSpecialColumn Then
            vFieldName = mvEnv.Connection.DBSpecialCol(vPrefix, vAttrName)
          Else
            vFieldName = String.Format("{0}.{1}", vPrefix, vAttrName)
          End If
          pWhereFields.Add(vFieldName, vFieldType, vValue, vOperator)
          If vOperator = CDBField.FieldWhereOperators.fwoBetweenFrom Then
            pWhereFields.Add(vFieldName & "#2", vFieldType, vBetweenValues(1), CDBField.FieldWhereOperators.fwoBetweenTo)
          End If
          If vRowNumber > pRowCount Then pRowCount = vRowNumber
          vCount += 1
        End If
      Next

      'BR20313
      If vDefaultAddressOnly AndAlso vPrefixRowZeroExists Then
        If pPageCode = "QECA" Then
          pWhereFields.Add("c.address_number", CDBField.FieldTypes.cftInteger, "a0.address_number")
        ElseIf pPageCode = "QEOA" Then
          pWhereFields.Add("o.address_number", CDBField.FieldTypes.cftInteger, "a0.address_number")
        End If
      End If

      Return vCount
    End Function

    Private Sub AddFieldsFromList(ByVal pFields As StringBuilder, ByVal pList As String)
      Dim vFields As New StringList(pFields.ToString)
      Dim vItems As New StringList(pList)
      For Each vItem As String In vItems
        Dim vPos As Integer = vItem.IndexOf(".")
        Dim vFieldName As String
        If vPos > 0 Then
          vFieldName = vItem.Substring(vPos + 1)
          If vFieldName = "address_number" Then
            If pFields.ToString.Contains(".address_number") Then   'Special case from address number coming from multiple tables
              Continue For
            End If
          End If
        Else
          vFieldName = vItem
        End If
        If Not (vFields.Contains(vFieldName) OrElse vFields.Contains(vItem)) Then
          pFields.Append(",")
          pFields.Append(vItem)
        End If
      Next
    End Sub

    Private Function FieldUsesPrefix(ByVal pWhereFields As CDBFields, ByVal pPrefix As String) As Boolean
      For Each vField As CDBField In pWhereFields
        If vField.Name.StartsWith(pPrefix & ".") Then Return True
      Next
    End Function

    Public Sub ProcessPageParameters(ByVal pPageCode As String, ByVal pWhereFields As CDBFields, ByVal pAnsiJoins As AnsiJoins, ByVal pExcludes As List(Of SQLStatement))
      Dim vQBETables As New List(Of QBETable)
      Dim vAlias As String = ""
      Dim vQBEType As QBETypes = QBETypes.Contacts

      Select Case pPageCode
        Case "QECO"                     'Contacts
          vAlias = "c"
          vQBETables.Add(New QBETable("principal_users", "pu", "PrincipalUser,PrincipalUserReason"))
          vQBETables.Add(New QBETable("vat_registration_numbers", "vn", "VatNumber"))
        Case "QECA"                     'Addresses
          vAlias = "a"
          vQBETables.Add(New QBETable("contact_addresses", "ca", "Current,ValidFrom,ValidTo"))
          vQBETables.Add(New QBETable("addresses", "a"))
        Case "QECN"                     'Numbers
          vAlias = "com"
          vQBETables.Add(New QBETable("communications", "com"))
        Case "QECC"                     'Categories
          vAlias = "cc"
          vQBETables.Add(New QBETable("contact_categories", "cc"))
        Case "QECM"                     'Members
          vAlias = "m"
          vQBETables.Add(New QBETable("members", "m"))
        Case "QECP"                     'Payment Plans
          vAlias = "pp"
          vQBETables.Add(New QBETable("orders", "pp"))
        Case "QECJ"                     'Positions
          vAlias = "cp"
          vQBETables.Add(New QBETable("organisations", "po", "Name"))
          vQBETables.Add(New QBETable("contact_positions", "cp"))
        Case "QECR"                     'Roles
          vAlias = "cr"
          vQBETables.Add(New QBETable("contact_roles", "cr"))
        Case "QECL"                     'Mailings
          vAlias = "cm"
          vQBETables.Add(New QBETable("mailing_history", "mh", "Mailing,MailingDate,MailingBy,Topic,SubTopic,DocumentSubject"))
          vQBETables.Add(New QBETable("contact_mailings", "cm"))
        Case "QECS"                     'Suppressions
          vAlias = "cs"
          vQBETables.Add(New QBETable("contact_suppressions", "cs"))
        Case "QECD"                     'Documents
          vAlias = "cl"
          vQBETables.Add(New QBETable("communications_log_subjects", "cls", "Topic,SubTopic,Quantity"))
          vQBETables.Add(New QBETable("communications_log", "cl"))
        Case "QECF"                     'Financial
          vAlias = "bt"
          vQBETables.Add(New QBETable("batches", "b", "BatchType,PayingInSlipNumber"))
          vQBETables.Add(New QBETable("batch_transactions", "bt"))
        Case "QECE"                     'Events
          vAlias = "ev"
          vQBETables.Add(New QBETable("sessions", "s", "Subject,SkillLevel"))
          vQBETables.Add(New QBETable("event_topics", "et", "Topic,SubTopic"))
          vQBETables.Add(New QBETable("events", "ev"))
        Case "QECX"
          vQBETables.Add(New QBETable("exam_units", "eu", "ExamUnitCode"))
          vQBETables.Add(New QBETable("exam_sessions", "es", "ExamSessionCode"))
          vQBETables.Add(New QBETable("exam_centres", "ec", "ExamCentreCode"))
          vQBETables.Add(New QBETable("exam_bookings", "eb", "ExamUnitId,ExamCentreId,ExamSessionId"))
          vQBETables.Add(New QBETable("exam_booking_units", "ebu", "ExamCandidateNumber,ExamStudentUnitStatus,TotalMark,TotalGrade,TotalResult"))
          vQBETables.Add(New QBETable("exam_schedule", "esc", "StartDate"))
          vQBETables.Add(New QBETable("exam_booking_units", "ebu"))
          vQBETables.Add(New QBETable("contact_exam_certs", "ecr", "ExamCertNumberPrefix,ExamCertNumber,ExamCertNumberSuffix"))
        Case "QECH"                     'Relationships From
          vAlias = "rf"
          vQBETables.Add(New QBETable("contact_links", "rf"))
        Case "QECT"                     'Relationships To
          vAlias = "rt"
          vQBETables.Add(New QBETable("contact_links", "rt"))
        Case "QEOO"                     'Organisations
          vQBEType = QBETypes.Organisations
          vAlias = "o"
          vQBETables.Add(New QBETable("contacts", "dc", "LabelName,Salutation,VatCategory"))
          vQBETables.Add(New QBETable("principal_user", "pu", "PrincipalUser,PrincipalUserReason"))
          vQBETables.Add(New QBETable("vat_registration_numbers", "vn", "VatNumber"))
        Case "QEOA"                     'Addresses
          vQBEType = QBETypes.Organisations
          vAlias = "a"
          vQBETables.Add(New QBETable("organisation_addresses", "oa", "Current,ValidFrom,ValidTo"))
          vQBETables.Add(New QBETable("addresses", "a"))
        Case "QEON"                     'Numbers
          vQBEType = QBETypes.Organisations
          vAlias = "com"
          vQBETables.Add(New QBETable("communications", "com"))
        Case "QEOC"                     'Categories
          vQBEType = QBETypes.Organisations
          vAlias = "oc"
          vQBETables.Add(New QBETable("organisation_categories", "oc"))
        Case "QEOM"                     'Members
          vQBEType = QBETypes.Organisations
          vAlias = "m"
          vQBETables.Add(New QBETable("members", "m"))
        Case "QEOP"                     'Payment Plans
          vQBEType = QBETypes.Organisations
          vAlias = "pp"
          vQBETables.Add(New QBETable("orders", "pp"))
        Case "QEOJ"                     'Positions
          vQBEType = QBETypes.Organisations
          vAlias = "cp"
          vQBETables.Add(New QBETable("contact_positions", "cp"))
        Case "QEOR"                     'Roles
          vQBEType = QBETypes.Organisations
          vAlias = "cr"
          vQBETables.Add(New QBETable("contact_roles", "cr"))
        Case "QEOL"                     'Mailings
          vQBEType = QBETypes.Organisations
          vAlias = "cm"
          vQBETables.Add(New QBETable("mailing_history", "mh", "Mailing,MailingDate,MailingBy,Topic,SubTopic,DocumentSubject"))
          vQBETables.Add(New QBETable("contact_mailings", "cm"))
        Case "QEOS"                     'Suppressions
          vQBEType = QBETypes.Organisations
          vAlias = "cs"
          vQBETables.Add(New QBETable("organisation_suppressions", "cs"))
        Case "QEOD"                     'Documents
          vQBEType = QBETypes.Organisations
          vAlias = "cl"
          vQBETables.Add(New QBETable("communications_log_subjects", "cls", "Topic,SubTopic,Quantity"))
          vQBETables.Add(New QBETable("communications_log", "cl"))
        Case "QEOF"                     'Financial
          vQBEType = QBETypes.Organisations
          vAlias = "bt"
          vQBETables.Add(New QBETable("batches", "b", "BatchType,PayingInSlipNumber"))
          vQBETables.Add(New QBETable("batch_transactions", "bt"))
        Case "QEOE"                     'Events
          vQBEType = QBETypes.Organisations
          vAlias = "ev"
          vQBETables.Add(New QBETable("sessions", "s", "Subject,SkillLevel"))
          vQBETables.Add(New QBETable("event_topics", "et", "Topic,SubTopic"))
          vQBETables.Add(New QBETable("events", "ev"))
        Case "QEOH"                     'Relationships From
          vQBEType = QBETypes.Organisations
          vAlias = "rf"
          vQBETables.Add(New QBETable("organisation_links", "rf"))
        Case "QEOT"                     'Relationships To
          vQBEType = QBETypes.Organisations
          vAlias = "rt"
          vQBETables.Add(New QBETable("organisation_links", "rt"))
        Case "QEEH"                     'Event Header
          vQBEType = QBETypes.Events
          vAlias = "e"
          vQBETables.Add(New QBETable("sessions", "ds", "Subject,SkillLevel,StartTime,EndDate,EndTime,Location,Notes"))
        Case "QEED"                     'Event details
          vQBEType = QBETypes.Events
          vAlias = "e"
          vQBETables.Add(New QBETable("sessions", "ds", "MaximumAttendees,MinimumAttendees,TargetAttendees,NumberInterested,NumberOfAttendees,NumberOnWaitingList,MaximumOnWaitingList"))
        Case "QEEO"                     'Event organiser
          vQBEType = QBETypes.Events
          vAlias = "e"
          vQBETables.Add(New QBETable("event_organisers", "eo", "Organiser,OrganiserReference,PriceToAttendees,Product,Rate"))
        Case "QEES"                     'Event sessions
          vQBEType = QBETypes.Events
          vAlias = "s"
          vQBETables.Add(New QBETable("sessions", "s"))
        Case "QEEV"                     'Event venues
          vQBEType = QBETypes.Events
          vAlias = "vb"
          vQBETables.Add(New QBETable("event_venue_bookings", "vb"))
        Case "QEET"                     'Event topics
          vQBEType = QBETypes.Events
          vAlias = "et"
          vQBETables.Add(New QBETable("event_topics", "et"))
        Case "QECU"                     'Fundraising Requests
          vAlias = "fr"
          vQBETables.Add(New QBETable("fundraising_requests", "fr"))
        Case "QEOU"                     'Fundraising Requests
          vQBEType = QBETypes.Organisations
          vAlias = "fr"
          vQBETables.Add(New QBETable("fundraising_requests", "fr"))
      End Select


      Dim vAlternates As New CDBParameters
      For Each vQBETable As QBETable In vQBETables
        If vQBETable.Attributes.Length > 0 Then vAlternates.Add(vQBETable.AliasName, vQBETable.Attributes)
      Next

      Dim vMaxRows As Integer = 0
      If AddPageParameters(pPageCode, QBERowTypes.Include, pWhereFields, vAlias, vAlternates, vMaxRows) > 0 Then
        'There are some parameters from this page
        For Each vQBETable As QBETable In vQBETables
          Dim vIndex As Integer = 0
          Dim vPrefix As String = vQBETable.AliasName & vIndex.ToString
          While vIndex <= vMaxRows
            If FieldUsesPrefix(pWhereFields, vPrefix) Then
              If (pPageCode = "QEEH" OrElse pPageCode = "QEED") AndAlso vQBETable.AliasName = "ds0" AndAlso _
                Not pWhereFields.ContainsKey(vPrefix & ".session_type") Then
                pWhereFields.Add(vPrefix & ".session_type", "0")
              End If
              If pPageCode = "QECM" AndAlso pWhereFields.ContainsKey(vPrefix & ".corporate") Then
                Dim vCorporate As String = vPrefix & ".corporate"
                If pWhereFields(vCorporate).Value = "Y" Then
                  vQBETable.Corporate = True
                  pWhereFields.Add(vPrefix.Replace("m", "cp") & ".current", "Y").SpecialColumn = True
                End If
                pWhereFields.Remove(vCorporate)
              End If

              vQBETable.AddAnsiJoins(mvEnv, vQBEType, pAnsiJoins, vIndex)
            End If
            vIndex += 1
            vPrefix = vQBETable.AliasName & vIndex.ToString
          End While
        Next
      End If
      'Session Type O
      If pPageCode = "QECA" AndAlso Not pAnsiJoins.ContainsAnyJoinToTable("addresses") Then
        pAnsiJoins.Add("addresses a", "c.address_number", "a.address_number")
      ElseIf pPageCode = "QEOA" AndAlso Not pAnsiJoins.ContainsAnyJoinToTable("addresses") Then
        pAnsiJoins.Add("addresses a", "o.address_number", "a.address_number")
      End If

      'Now handle excludes?
      vMaxRows = 0
      Dim vWhereFields As New CDBFields
      Dim vAnsiJoins As New AnsiJoins
      If AddPageParameters(pPageCode, QBERowTypes.Exclude, vWhereFields, vAlias, vAlternates, vMaxRows) > 0 Then
        'There are some parameters from this page
        For Each vQBETable As QBETable In vQBETables
          Dim vIndex As Integer = 0
          Dim vPrefix As String = vQBETable.AliasName & vIndex.ToString
          While vIndex <= vMaxRows
            If FieldUsesPrefix(vWhereFields, vPrefix) Then
              If (pPageCode = "QEEH" OrElse pPageCode = "QEED") AndAlso vQBETable.AliasName = "ds" Then
                vWhereFields.Add(vPrefix & ".session_type", "0")
              End If
              vQBETable.AddAnsiJoins(mvEnv, vQBEType, vAnsiJoins, vIndex)
            End If
            vIndex += 1
            vPrefix = vQBETable.AliasName & vIndex.ToString
          End While
        Next
        Dim vPrimaryAttribute As String = ""
        Dim vLinkedAttribute As String = ""
        Dim vPrimaryTable As String = ""
        Select Case vQBEType
          Case QBETypes.Contacts
            vPrimaryAttribute = "c.contact_number"
          Case QBETypes.Organisations
            vPrimaryAttribute = "o.organisation_number"
          Case QBETypes.Events
            vPrimaryAttribute = "e.event_number"
        End Select
        For Each vJoin As AnsiJoin In vAnsiJoins
          If vJoin.Joins(0).Attribute2 = vPrimaryAttribute Then
            vPrimaryTable = vJoin.TableName
            vLinkedAttribute = vJoin.Joins(0).Attribute1
            vAnsiJoins.Remove(vJoin)
            Exit For
          End If
        Next
        vWhereFields.AddJoin(vLinkedAttribute, vPrimaryAttribute)
        pExcludes.Add(New SQLStatement(mvEnv.Connection, "*", vPrimaryTable, vWhereFields, "", vAnsiJoins))
      End If
    End Sub

    Private Class QBETable
      Private mvTableName As String
      Private mvAlias As String
      Private mvAttributes As String
      Private mvCorporate As Boolean

      Public Sub New(ByVal pTableName As String, ByVal pAlias As String)
        mvTableName = pTableName
        mvAlias = pAlias
        mvAttributes = ""
      End Sub

      Public Sub New(ByVal pTableName As String, ByVal pAlias As String, ByVal pAttributes As String)
        mvTableName = pTableName
        mvAlias = pAlias
        mvAttributes = pAttributes
      End Sub

      Public Property Corporate As Boolean
        Get
          Return mvCorporate
        End Get
        Set(pValue As Boolean)
          mvCorporate = pValue
        End Set
      End Property


      Public ReadOnly Property TableName() As String
        Get
          Return mvTableName
        End Get
      End Property

      Public ReadOnly Property Attributes() As String
        Get
          Return mvAttributes
        End Get
      End Property

      Public ReadOnly Property AliasName() As String
        Get
          Return mvAlias
        End Get
      End Property

      Public Sub AddAnsiJoins(pEnv As CDBEnvironment, ByVal pQBEtype As QBETypes, ByVal pAnsiJoins As AnsiJoins, ByVal pIndex As Integer)
        Dim vTableName As String = mvTableName
        If vTableName = "addresses" Then
          Select Case pQBEtype
            Case QBETypes.Contacts
              vTableName = "contact_addresses"
            Case QBETypes.Organisations
              vTableName = "organisation_addresses"
          End Select
        End If
        Select Case vTableName
          Case "contact_mailings"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, GetMailingsUnion(pEnv), "cm", "contact_number", "c", "contact_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, GetMailingsUnion(pEnv), "cm", "contact_number", "o", "organisation_number", pIndex)
            End Select
          Case "mailing_history"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, GetMailingsUnion(pEnv), "cm", "contact_number", "c", "contact_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, GetMailingsUnion(pEnv), "cm", "contact_number", "o", "organisation_number", pIndex)
            End Select
            AddAnsiJoin(pAnsiJoins, "mailing_history mh", "mh", "mailing_number", "cm", "mailing_number", pIndex)
          Case "batches"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, "batch_transactions bt", "bt", "contact_number", "c", "contact_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, "batch_transactions bt", "bt", "contact_number", "o", "organisation_number", pIndex)
            End Select
            AddAnsiJoin(pAnsiJoins, "batches b", "b", "batch_number", "bt", "batch_number", pIndex)
          Case "communications"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "contact_number", "c", "contact_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, "organisation_addresses oa", "oa", "organisation_number", "o", "organisation_number", pIndex)
                AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "address_number", "oa", "address_number", pIndex)
            End Select
          Case "communications_log"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, "communications_log_links cll", "cll", "contact_number", "c", "contact_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, "organisation_addresses oa", "oa", "organisation_number", "o", "organisation_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "communications_log_links cll", "cll", "address_number", "oa", "address_number", pIndex)
            End Select
            AddAnsiJoin(pAnsiJoins, "communications_log cl", "cl", "communications_log_number", "cll", "communications_log_number", pIndex)
          Case "communications_log_subjects"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, "communications_log_links cll", "cll", "contact_number", "c", "contact_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, "organisation_addresses oa", "oa", "organisation_number", "o", "organisation_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "communications_log_links cll", "cll", "address_number", "oa", "address_number", pIndex)
            End Select
            AddAnsiJoin(pAnsiJoins, "communications_log cl", "cl", "communications_log_number", "cll", "communications_log_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "communications_log_subjects cls", "cls", "communications_log_number", "cl", "communications_log_number", pIndex)
          Case "contact_addresses"
            AddAnsiJoin(pAnsiJoins, "contact_addresses ca", "ca", "contact_number", "c", "contact_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "addresses a", "a", "address_number", "ca", "address_number", pIndex)
          Case "organisation_addresses"
            AddAnsiJoin(pAnsiJoins, "organisation_addresses oa", "oa", "organisation_number", "o", "organisation_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "addresses a", "a", "address_number", "oa", "address_number", pIndex)
          Case "sessions"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, "delegates d", "d", "contact_number", "c", "contact_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "sessions s", "s", "event_number", "d", "event_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, "event_bookings d", "d", "contact_number", "o", "organisation_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "sessions s", "s", "event_number", "d", "event_number", pIndex)
              Case QBETypes.Events
                AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "event_number", "e", "event_number", pIndex)
            End Select
          Case "events"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, "delegates d", "d", "contact_number", "c", "contact_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, "event_bookings d", "d", "contact_number", "o", "organisation_number", pIndex)
            End Select
            AddAnsiJoin(pAnsiJoins, "events ev", "ev", "event_number", "d", "event_number", pIndex)
          Case "event_topics"
            Select Case pQBEtype
              Case QBETypes.Contacts
                AddAnsiJoin(pAnsiJoins, "delegates d", "d", "contact_number", "c", "contact_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "events ev", "ev", "event_number", "d", "event_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "event_topics et", "et", "event_number", "ev", "event_number", pIndex)
              Case QBETypes.Organisations
                AddAnsiJoin(pAnsiJoins, "event_bookings d", "d", "contact_number", "o", "organisation_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "events ev", "ev", "event_number", "d", "event_number", pIndex)
                AddAnsiJoin(pAnsiJoins, "event_topics et", "et", "event_number", "ev", "event_number", pIndex)
              Case QBETypes.Events
                AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "event_number", "e", "event_number", pIndex)
            End Select
          Case "exam_centres"
            AddAnsiJoin(pAnsiJoins, "exam_bookings eb", "eb", "contact_number", "c", "contact_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "exam_centres ec", "ec", "exam_centre_id", "eb", "exam_centre_id", pIndex)
          Case "exam_schedule"
            AddAnsiJoin(pAnsiJoins, "exam_booking_units ebu", "ebu", "contact_number", "c", "contact_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "exam_schedule esc", "esc", "exam_schedule_id", "ebu", "exam_schedule_id", pIndex)
          Case "exam_sessions"
            AddAnsiJoin(pAnsiJoins, "exam_bookings eb", "eb", "contact_number", "c", "contact_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "exam_sessions es", "es", "exam_session_id", "eb", "exam_session_id", pIndex)
          Case "exam_units"
            AddAnsiJoin(pAnsiJoins, "exam_bookings eb", "eb", "contact_number", "c", "contact_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "exam_booking_units ebu", "ebu", "exam_booking_id", "eb", "exam_booking_id", pIndex)
            AddAnsiJoin(pAnsiJoins, "exam_units eu", "eu", "exam_unit_id", "ebu", "exam_unit_id", pIndex)
          Case "contact_exam_certs"
            AddAnsiJoin(pAnsiJoins, "exam_bookings eb", "eb", "contact_number", "c", "contact_number", pIndex)
            AddAnsiJoin(pAnsiJoins, "exam_booking_units ebu", "ebu", "exam_booking_id", "eb", "exam_booking_id", pIndex)
            AddAnsiJoin(pAnsiJoins, "contact_exam_certs ecr", "ecr", "exam_booking_unit_id", "ebu", "exam_booking_unit_id", pIndex)
          Case "contact_links"
            If AliasName = "rf" Then
              'union is used because relationships might be in organisation links or contact links depending on who the relationship is from
              AddAnsiJoin(pAnsiJoins, GetContactOrgLinksUnion(pEnv), "rf", "number2", "c", "contact_number", pIndex)
            Else
              AddAnsiJoin(pAnsiJoins, "contact_links rt", "rt", "contact_number_1", "c", "contact_number", pIndex)
            End If
          Case "organisation_links"
            If AliasName = "rf" Then
              'union is used because relationships might be in organisation links or contact links depending on who the relationship is from
              AddAnsiJoin(pAnsiJoins, GetContactOrgLinksUnion(pEnv), "rf", "number2", "o", "organisation_number", pIndex)
            Else
              AddAnsiJoin(pAnsiJoins, "organisation_links rt", "rt", "organisation_number_1", "o", "organisation_number", pIndex)
            End If
          Case Else
            Select Case pQBEtype
              Case QBETypes.Contacts
                Select Case vTableName
                  Case "organisations"
                    AddAnsiJoin(pAnsiJoins, "contact_positions cp", "cp", "contact_number", "c", "contact_number", pIndex)
                    AddAnsiJoin(pAnsiJoins, "organisations po", "po", "organisation_number", "cp", "organisation_number", pIndex)
                  Case "members"
                    If mvCorporate Then
                      AddAnsiJoin(pAnsiJoins, "contact_positions cp", "cp", "contact_number", "c", "contact_number", pIndex)
                      AddAnsiJoin(pAnsiJoins, "organisations po", "po", "organisation_number", "cp", "organisation_number", pIndex)
                      AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "contact_number", "po", "organisation_number", pIndex)
                    Else
                      AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "contact_number", "c", "contact_number", pIndex)
                    End If
                  Case Else
                    AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "contact_number", "c", "contact_number", pIndex)
                End Select
              Case QBETypes.Organisations
                Select Case vTableName
                  Case "communications"
                    AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "address_number", "a", "address_number", pIndex)
                  Case "members", "orders", "batch_transactions", "contacts", "fundraising_requests" 'BR19022 added QBE fundraising requests
                    AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "contact_number", "o", "organisation_number", pIndex)
                  Case Else
                    AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "organisation_number", "o", "organisation_number", pIndex)
                End Select
              Case QBETypes.Events
                AddAnsiJoin(pAnsiJoins, String.Format("{0} {1}", mvTableName, mvAlias), mvAlias, "event_number", "e", "event_number", pIndex)
            End Select
        End Select
      End Sub

      Private Sub AddAnsiJoin(ByVal pAnsiJoins As AnsiJoins, ByVal pTableNameAndAlias As String, ByVal pFromAlias As String, ByVal pFromAttribute As String, ByVal pToAlias As String, ByVal pToAttribute As String, ByVal pIndex As Integer)
        If Not pAnsiJoins.ContainsJoinToTable(pTableNameAndAlias & pIndex.ToString) Then
          Dim vFromAtttribute As String = String.Format("{0}{1}.{2}", pFromAlias, pIndex.ToString, pFromAttribute)
          Dim vToAtttribute As String
          If pToAlias = "c" OrElse pToAlias = "o" OrElse pToAlias = "e" Then
            vToAtttribute = String.Format("{0}.{1}", pToAlias, pToAttribute)
          Else
            vToAtttribute = String.Format("{0}{1}.{2}", pToAlias, pIndex.ToString, pToAttribute)
          End If
          pAnsiJoins.Add(pTableNameAndAlias & pIndex.ToString, vFromAtttribute, vToAtttribute)
        End If
      End Sub

      Private Function GetMailingsUnion(pEnv As CDBEnvironment) As String
        Dim vSQL1 As New SQLStatement(pEnv.Connection, "contact_number, mailing_number", "contact_mailings", New CDBFields)
        vSQL1.AddUnion(New SQLStatement(pEnv.Connection, "contact_number, mailing_number", "contact_emailings", New CDBFields))
        Return String.Format("( {0} ) cm", vSQL1.SQL)
      End Function

      Private Function GetContactOrgLinksUnion(pEnv As CDBEnvironment) As String
        Dim vSQL1 As New SQLStatement(pEnv.Connection, "contact_number_1 AS number1, contact_number_2 AS number2, relationship, valid_from, valid_to, relationship_status, historical," & pEnv.Connection.DBMaxToString("notes") & "AS notes", "contact_links", New CDBFields)
        vSQL1.AddUnion(New SQLStatement(pEnv.Connection, "organisation_number_1 AS number1, organisation_number_2 AS number2, relationship, valid_from, valid_to, relationship_status, historical," & pEnv.Connection.DBMaxToString("notes") & "AS notes", "organisation_links", New CDBFields))
        Return String.Format("( {0} ) rf", vSQL1.SQL)
      End Function
    End Class
  End Class
End Namespace

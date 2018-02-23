Namespace Access
  Public Class MarketingData

    Private mvLogFile As LogFile

    Private Class Performance
      Public Performance As String = ""
      Public LowerLevel As Double
      Public UpperLevel As Double
      Public RollingBoundary As Integer
      Public IncludeBefore As String = ""
      Public IncludeAfter As String = ""
      Public AllPayMethods As String = ""
      Public PaymentMethods() As String
      Public MailingsBefore As String = ""
      Public MailingsAfter As String = ""
      Public ExpenditureGroup As String = ""
      Public ProcessMailings As Boolean
      Public IncludeJoint As Boolean
      Public ContactNumber As Integer
      Public NoPayments As Integer
      Public NumberAbove As Integer
      Public NumberBelow As Integer
      Public NumberBetween As Integer
      Public NoMailings As Integer
      Public NoResponse As Integer
      Public ValuePayments As Double
      Public ValueAbove As Double
      Public ValueBelow As Double
      Public ValueBetween As Double
      Public FirstPayment As Double
      Public LastPayment As Double
      Public MaximumPayment As Double
      Public RollingValue As Double
      Public PrevRollingValue As Double
      Public AveragePerMailing As Double
      Public AverageValue As Double
      Public ResponseRate As Double
      Public FirstPaymentDate As String = ""
      Public LastPaymentDate As String = ""
      Public MaximumPaymentDate As String = ""
      Public RollingDateFrom As Date
      Public RollingDateTo As Date
      Public PrevRollingDateFrom As Date
      Public PrevRollingDateTo As Date
    End Class

    Private Class ScoringRow
      Public Score As String
      Public SearchAreaName As String = ""
      Public ContactHeader As Boolean
      Public ContactAttr As String = ""
      Public CorO As String = ""
      Public IorE As String = ""
      Public MainValue As String = ""
      Public SubsidiaryValue As String = ""
      Public PeriodValue As String = ""
      Public Multiplier As Double
      Public Points As Double
      Public ContactPoints As Double
      Public NumericValue As Double
      Public ValueRequired As Boolean
      Public TableName As String = ""
      Public FirstSpecial As String = ""
      Public SpecialTables As String = ""
      Public SpecialLink As String = ""
      Public MainAttribute As String = ""
      Public MainDataType As String = ""
      Public SubsidiaryAttribute As String = ""
      Public SubsidiaryDataType As String = ""
      Public FromAttribute As String = ""
      Public ToAttribute As String = ""
      Public MainValueTo As String = ""
      Public SubsidiaryValueTo As String = ""
      Public PeriodValueTo As String = ""
      Public RecordSet As CDBRecordSet
      Public EndOfRecordSet As Boolean
    End Class

    Public Function ProcessHeader(ByVal pEnv As CDBEnvironment, ByVal pConn As CDBConnection, ByRef pJob As JobSchedule, ByVal pContactNumber As Integer, Optional ByVal pUseOwner As Boolean = False, Optional ByVal pNewContacts As Boolean = False, Optional ByRef pIncludeSet As Integer = 0, Optional ByRef pExcludeSet As Integer = 0) As Integer
      Dim vWhere As String = ""
      Dim vYears As Integer
      Dim vStartDate As String
      Dim vHeaderRecords As Integer
      Dim vWhereFields As New CDBFields
      Dim vInsertFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vCDBIndexes As New CDBIndexes
      Dim vIndexesDropped As Boolean
      Dim vLastContactNo As Integer

      Try
        If pNewContacts Then vLastContactNo = CInt(pConn.GetValue("SELECT last_contact_number FROM marketing_controls"))

        If pContactNumber > 0 Then
          vWhere = "contact_number = " & pContactNumber
        Else
          If vLastContactNo > 0 Then
            vWhere = "contact_number > " & vLastContactNo
          Else
            LogStatus(pJob, (ProjectText.String30301)) 'Dropping Indexes for Contact Header
            vCDBIndexes.Init(pConn, "contact_header")
            vCDBIndexes.DropAll(pConn)
            vIndexesDropped = True
          End If
        End If

        LogStatus(pJob, (ProjectText.String30302)) 'Clearing Contact Header Data
        If pContactNumber > 0 Then
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
          pConn.DeleteRecords("contact_header", vWhereFields, False)
        Else
          If vLastContactNo > 0 Then
            vWhereFields.Add("contact_number", vLastContactNo, CDBField.FieldWhereOperators.fwoGreaterThan)
            pConn.DeleteRecords("contact_header", vWhereFields, False)
          Else
            pConn.DeleteAllRecords("contact_header")
          End If
        End If

        LogStatus(pJob, (ProjectText.String30303)) 'Creating Contact Header Data
        vYears = IntegerValue(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMinimumAge)) * -1
        vStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, vYears, Today))
        Dim vSQL As New StringBuilder
        vSQL.Append("INSERT INTO contact_header (contact_number,address_number,sex,source,source_date,status,status_date,age,")
        vSQL.Append("postal_area,postal_district,branch,country,number_of_bankers_orders,value_of_bankers_orders,")
        vSQL.Append("number_of_direct_debits, value_of_direct_debits, number_of_covenants, value_of_covenants")
        If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHeaderPostalSector) Then vSQL.Append(", postal_sector")
        If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataNumberofCCCAs) Then
          vSQL.Append(",number_of_cccas,value_of_cccas,number_of_non_apm,value_of_non_apm,total_number_of_pay_plans,total_value_of_pay_plans,number_of_pg_pledges,number_of_gads")
        End If
        vSQL.Append(") ")

        vSQL.Append("SELECT c.contact_number,c.address_number,sex,source,source_date,status,status_date,")
        vSQL.Append(pConn.DBIsNull(pConn.DBAge(), "999"))
        'The next is to get the postal area
        'First get the string to the left of any space and add 4 spaces to the back of it
        Dim vPAFirstPart As String = pConn.DBLeft("postcode", pConn.DBIndexOf("' '", "postcode"))
        Dim vPAResult2 As String = vPAFirstPart & pConn.DBConcatString & "'    '"
        'Then truncate it to 4 characters - at this point any postcodes with no spaces in will just have a string of 4 spaces
        Dim vPAResult3 As String = pConn.DBLeft(vPAResult2, "4")
        'Now add on the first 4 characters of the postcode
        Dim vPAResult4 As String = vPAResult3 & pConn.DBConcatString & pConn.DBLeft("postcode", "4")
        'The trim off the leading spaces this will move the 4 chars for those with no spaces to the left of the string
        Dim vPAResult5 As String = pConn.DBLTrim(vPAResult4)
        'The just truncate to 4 characters and remove trailing spaces - easy ain't it....
        Dim vPAResult6 As String = pConn.DBRTrim(pConn.DBLeft(vPAResult5, "4"))

        vSQL.Append(",")
        vSQL.Append(vPAResult6)
        vSQL.Append(",")
        vSQL.Append(pConn.DBLeft("postcode", "2"))
        vSQL.Append(",branch,country,")
        vSQL.Append(pConn.DBIsNull("number_of_bankers_orders", "0"))
        vSQL.Append(",")
        vSQL.Append(pConn.DBIsNull("value_of_bankers_orders", "0"))
        vSQL.Append(",")
        vSQL.Append(pConn.DBIsNull("number_of_direct_debits", "0"))
        vSQL.Append(",")
        vSQL.Append(pConn.DBIsNull("value_of_direct_debits", "0"))
        vSQL.Append(",")
        vSQL.Append(pConn.DBIsNull("number_of_covenants", "0"))
        vSQL.Append(",")
        vSQL.Append(pConn.DBIsNull("value_of_covenants", "0"))
        If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataHeaderPostalSector) Then
          'This is to get the postal sector
          'get the first part of postcode before the space and concat to the first char after the space
          Dim vPSFirstPart As String = pConn.DBRTrim(pConn.DBLeft("postcode", pConn.DBIndexOf("' '", "postcode")))
          Dim vPSFirstAfterSpace As String = pConn.DBLeft(pConn.DBSubString("postcode", pConn.DBIndexOf("' '", "postcode") & "+ 1", pConn.DBIndexOf("' '", "postcode")), "1")
          Dim vPSResult1 As String = vPSFirstPart & pConn.DBConcatString & vPSFirstAfterSpace
          'Now add 5 spaces on to the end of it
          Dim vPSResult2 As String = vPSResult1 & pConn.DBConcatString & "'     '"
          'Then truncate it to 5 characters - at this point any postcodes with no spaces in will just have a string of 5 spaces
          Dim vPSResult3 As String = pConn.DBLeft(vPSResult2, "5")
          'Now add on the first 5 characters of the postcode
          Dim vPSResult4 As String = vPSResult3 & pConn.DBConcatString & pConn.DBLeft("postcode", "5")
          'The trim off the leading spaces this will move the 5 chars for those with no spaces to the left of the string
          Dim vPSResult5 As String = pConn.DBLTrim(vPSResult4)
          'The just truncate to 5 characters and remove trailing spaces - easy ain't it....
          Dim vPSResult6 As String = pConn.DBRTrim(pConn.DBLeft(vPSResult5, "5"))
          vSQL.Append(",")
          vSQL.Append(vPSResult6)
        End If
        If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataNumberofCCCAs) Then
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("number_of_cccas", "0"))
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("value_of_cccas", "0"))
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("number_of_non_apm", "0"))
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("value_of_non_apm", "0"))
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("total_number_of_pay_plans", "0"))
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("total_value_of_pay_plans", "0"))
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("number_of_pg_pledges", "0"))
          vSQL.Append(",")
          vSQL.Append(pConn.DBIsNull("number_of_gads", "0"))
        End If
        If pIncludeSet > 0 Then
          vSQL.Append(" FROM selected_contacts sc INNER JOIN contacts c ON sc.contact_number = c.contact_number")
        Else
          vSQL.Append(" FROM contacts c")
        End If
        If pUseOwner Then
          If pEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
            vSQL.Append(" INNER JOIN ownership_groups og ON c.ownership_group = og.ownership_group")
          Else
            vSQL.Append(" INNER JOIN contact_users cu ON c.contact_number = cu.contact_number")
          End If
        End If
        vSQL.Append(" INNER JOIN contact_addresses ca ON c.contact_number = ca.contact_number AND c.address_number = ca.address_number")
        vSQL.Append(" INNER JOIN addresses a ON a.address_number = ca.address_number")

        vSQL.Append(" LEFT OUTER JOIN (SELECT bo.contact_number, count(*) number_of_bankers_orders, SUM(((bo.amount * 12) / interval)) value_of_bankers_orders")
        vSQL.Append(" FROM bankers_orders bo, orders o, payment_frequencies pf WHERE bo.cancellation_reason IS NULL AND o.order_number = bo.order_number AND pf.payment_frequency = o.payment_frequency GROUP BY bo.contact_number) so ON c.contact_number = so.contact_number")

        vSQL.Append(" LEFT OUTER JOIN (SELECT dd.contact_number, count(*) number_of_direct_debits, SUM(((frequency_amount) * 12 / interval)) value_of_direct_debits")
        vSQL.Append(" FROM direct_debits dd, orders o, payment_frequencies pf WHERE dd.cancellation_reason IS NULL AND o.order_number = dd.order_number AND pf.payment_frequency = o.payment_frequency GROUP BY dd.contact_number) dds ON c.contact_number = dds.contact_number")

        vSQL.Append(" LEFT OUTER JOIN (SELECT co.contact_number, count(*) number_of_covenants, SUM(((frequency_amount * 12) / interval)) value_of_covenants")
        vSQL.Append(" FROM covenants co, orders o, payment_frequencies pf WHERE co.cancellation_reason IS NULL AND o.order_number = co.order_number AND pf.payment_frequency = o.payment_frequency GROUP BY co.contact_number) cov ON c.contact_number = cov.contact_number")

        If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataNumberofCCCAs) Then
          vSQL.Append(" LEFT OUTER JOIN (SELECT cca.contact_number, count(*) number_of_cccas, SUM(((frequency_amount) * 12 / interval)) value_of_cccas")
          vSQL.Append(" FROM credit_card_authorities cca, orders o, payment_frequencies pf WHERE cca.cancellation_reason IS NULL AND o.order_number = cca.order_number AND pf.payment_frequency = o.payment_frequency GROUP BY cca.contact_number) cccas ON c.contact_number = cccas.contact_number")

          vSQL.Append(" LEFT OUTER JOIN (SELECT o.contact_number, count(*) number_of_non_apm, SUM(((frequency_amount) * 12 / interval)) value_of_non_apm")
          vSQL.Append(" FROM orders o, payment_frequencies pf WHERE o.cancellation_reason IS NULL AND bankers_order <> 'Y' AND direct_debit <> 'Y' AND credit_card <> 'Y' and covenant <> 'Y' AND pf.payment_frequency = o.payment_frequency GROUP BY o.contact_number) non_apm ON c.contact_number = non_apm.contact_number")

          vSQL.Append(" LEFT OUTER JOIN (SELECT o.contact_number, count(*) total_number_of_pay_plans, SUM(((frequency_amount) * 12 / interval)) total_value_of_pay_plans")
          vSQL.Append(" FROM orders o, payment_frequencies pf WHERE o.cancellation_reason IS NULL AND pf.payment_frequency = o.payment_frequency GROUP BY o.contact_number) pps ON c.contact_number = pps.contact_number")

          vSQL.Append(" LEFT OUTER JOIN (SELECT gp.contact_number, count(*) number_of_pg_pledges")
          vSQL.Append(" FROM gaye_pledges gp WHERE gp.cancellation_reason IS NULL GROUP BY gp.contact_number) pgps ON c.contact_number = pgps.contact_number")

          vSQL.Append(" LEFT OUTER JOIN (SELECT gad.contact_number, count(*) number_of_gads")
          vSQL.Append(" FROM gift_aid_declarations gad LEFT OUTER JOIN orders o ON gad.order_number = o.order_number WHERE gad.cancellation_reason IS NULL AND batch_number IS NULL AND transaction_number IS NULL AND o.cancellation_reason IS NULL GROUP BY gad.contact_number) gads ON c.contact_number = gads.contact_number")
        End If

        vSQL.Append(" WHERE ")
        If pIncludeSet > 0 Then
          vSQL.Append(" sc.selection_set = ")
          vSQL.Append(pIncludeSet)
          vSQL.Append(" AND ")
        End If
        If vWhere.Length > 0 Then
          vSQL.Append("c.")
          vSQL.Append(vWhere)
          vSQL.Append(" AND ")
        End If
        vSQL.Append(" (date_of_birth IS NULL OR date_of_birth")
        vSQL.Append(pConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, vStartDate))
        vSQL.Append(")")
        If pExcludeSet > 0 Then
          vSQL.Append(" AND c.contact_number NOT IN (SELECT ec.contact_number FROM selected_contacts ec WHERE selection_set = ")
          vSQL.Append(pExcludeSet)
          vSQL.Append(")")
        End If
        If pUseOwner Then
          If pEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
            vSQL.Append(" AND og.principal_department = '")
          Else
            vSQL.Append(" AND cu.department = '")
          End If
          vSQL.Append(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDepartment) & "'")
        End If
        If pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlIncludeHistoric) = "N" Then vSQL.Append(" AND ca.historical = 'N'")
        vHeaderRecords = pConn.ExecuteSQL(pConn.ProcessAnsiJoins(vSQL.ToString))
        If vHeaderRecords = 0 And pContactNumber > 0 Then RaiseError(DataAccessErrors.daeNotAMarketingContact) 'Contact is not a Marketing Contact
        vWhereFields.Clear()
        vWhereFields.Add("postal_district", CDBField.FieldTypes.cftCharacter, "_[0-9]", CDBField.FieldWhereOperators.fwoLike)
        vUpdateFields.Clear()
        vUpdateFields.Add("postal_district", CDBField.FieldTypes.cftLong, pConn.DBLeft("postal_district", "1"))
        pConn.UpdateRecords("contact_header", vUpdateFields, vWhereFields, False)
        If vIndexesDropped Then
          LogStatus(pJob, (ProjectText.String30305)) 'Creating Indexes for Contact Header
          vCDBIndexes.ReCreate(pConn)
          vCDBIndexes.CreateIfMissing(pConn, True, {"contact_number"})
        End If
        ProcessHeader = vHeaderRecords
        If pNewContacts Then
          vLastContactNo = CInt(pConn.GetValue("SELECT MAX(contact_number) FROM contact_header"))
          vUpdateFields = New CDBFields
          vWhereFields = New CDBFields
          vUpdateFields.Add("last_contact_number", CDBField.FieldTypes.cftLong, vLastContactNo)
          vWhereFields.Add("department", CDBField.FieldTypes.cftCharacter, pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDepartment))
          pConn.UpdateRecords("marketing_controls", vUpdateFields, vWhereFields, False)
        End If
      Catch vEx As Exception
        PreserveStackTrace(vEx)
        If vIndexesDropped Then
          vCDBIndexes.ReCreate(pConn)
          vCDBIndexes.CreateIfMissing(pConn, True, {"contact_number"})
        End If
        Throw vEx
      End Try
    End Function

    Public Function ProcessExpenditure(ByVal pEnv As CDBEnvironment, ByVal pConn As CDBConnection, ByRef pJob As JobSchedule, ByVal pContactNumber As Integer, Optional ByRef pNewBatches As Boolean = False) As Integer
      Dim vSQL As String
      Dim vWhere As String
      Dim vRecordCount As Integer
      Dim vYears As Integer
      Dim vStartDate As String
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vLastBatchNo As Integer
      Dim vMaxBatchNo As Integer
      Dim vCDBIndexes As New CDBIndexes

      Try
        If pNewBatches Then vLastBatchNo = CInt(pConn.GetValue("SELECT last_batch_number FROM marketing_controls"))
        If pContactNumber > 0 Then
          vWhere = "contact_number = " & pContactNumber
        Else
          If vLastBatchNo > 0 Then
            vWhere = "batch_number > " & vLastBatchNo
          Else
            LogStatus(pJob, (ProjectText.String30311)) 'Dropping Indexes for Contact Expenditure
            vCDBIndexes.Init(pConn, "contact_expenditure")
            vCDBIndexes.DropAll(pConn)
          End If
        End If

        LogStatus(pJob, (ProjectText.String30312)) 'Clearing Contact Expenditure Data
        If pContactNumber > 0 Then
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
          pConn.DeleteRecords("contact_expenditure", vWhereFields, False)
        Else
          If vLastBatchNo = 0 Then pConn.DeleteAllRecords("contact_expenditure")
        End If

        LogStatus(pJob, (ProjectText.String30313)) 'Creating Contact Expenditure Data
        vYears = IntegerValue(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlEarliestDonation)) * -1
        vStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, vYears, Today))

        vSQL = "INSERT INTO contact_expenditure (contact_number, transaction_date, payment_method, product, rate, amount, expenditure_group, source, distribution_code)"
        vSQL = vSQL & " SELECT fh.contact_number, transaction_date, payment_method, fhd.product, rate, fhd.amount, expenditure_group, fhd.source, fhd.distribution_code"
        vSQL = vSQL & " FROM financial_history fh, financial_history_details fhd, product_groups pg"
        If pContactNumber > 0 Then
          vSQL = vSQL & " WHERE contact_number = " & pContactNumber & " AND "
        Else
          vSQL = vSQL & ",contact_header ch WHERE"
        End If
        If vLastBatchNo > 0 Then vSQL = vSQL & " fh.batch_number > " & vLastBatchNo & " AND "
        vSQL = vSQL & " fh.transaction_date " & pConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vStartDate) & " AND fh.amount >= 0"
        vSQL = vSQL & " AND fhd.batch_number = fh.batch_number AND fhd.transaction_number = fh.transaction_number"
        vSQL = vSQL & " AND fhd.amount >= 0 AND fhd.status IS NULL AND pg.product = fhd.product"
        If pContactNumber <= 0 Then vSQL = vSQL & " AND ch.contact_number = fh.contact_number"
        vRecordCount = pConn.ExecuteSQL(vSQL)

        If pContactNumber = 0 And vLastBatchNo = 0 Then
          LogStatus(pJob, (ProjectText.String30314)) 'Creating Indexes for Contact Expenditure
          vCDBIndexes.ReCreate(pConn)
          vCDBIndexes.CreateIfMissing(pConn, False, {"contact_number"})
        End If

        If pContactNumber = 0 And vRecordCount > 0 Then
          vMaxBatchNo = CInt(pConn.GetValue("SELECT MAX(batch_number) FROM financial_history"))
          vUpdateFields.Clear()
          vUpdateFields.Add("last_batch_number", CDBField.FieldTypes.cftLong, vMaxBatchNo)
          vWhereFields.Clear()
          vWhereFields.Add("department", CDBField.FieldTypes.cftCharacter, pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDepartment))
          pConn.UpdateRecords("marketing_controls", vUpdateFields, vWhereFields, False)
        End If
        ProcessExpenditure = vRecordCount
      Catch vEx As Exception
        PreserveStackTrace(vEx)
        vCDBIndexes.ReCreate(pConn)
        vCDBIndexes.CreateIfMissing(pConn, False, {"contact_number"})
        Throw vEx
      End Try
    End Function

    Public Function ProcessMembership(ByVal pEnv As CDBEnvironment, ByVal pConn As CDBConnection, ByRef pJob As JobSchedule, ByVal pContactNumber As Integer, Optional ByVal pNewMembers As Boolean = False, Optional ByRef pIncludeSet As Integer = 0, Optional ByRef pExcludeSet As Integer = 0) As Integer
      Dim vInsertFields As New CDBFields
      Dim vMemberFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vLastMembershipNo As Integer
      Dim vMaxMembershipNo As Integer
      Dim vCount As Integer
      Dim vCDBIndexes As New CDBIndexes

      Try
        'New members no longer supported??
        If pNewMembers And False Then vLastMembershipNo = CInt(pConn.GetValue("SELECT last_membership_number FROM marketing_controls"))

        If pContactNumber > 0 Then
          '
        Else
          If vLastMembershipNo > 0 Then
            '
          Else
            LogStatus(pJob, (ProjectText.String30325)) 'Dropping Indexes for Membership Header
            vCDBIndexes.Init(pConn, "membership_header")
            vCDBIndexes.DropAll(pConn)
          End If
        End If

        LogStatus(pJob, (ProjectText.String30327)) 'Clearing Membership Header Data
        If pContactNumber > 0 Then
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
          pConn.DeleteRecords("membership_header", vWhereFields, False)
        Else
          If vLastMembershipNo > 0 Then
            'Cannot delete records as we don't know what the membership number relates to in the table
          Else
            pConn.DeleteAllRecords("membership_header")
          End If
        End If

        LogStatus(pJob, (ProjectText.String30335)) 'Selecting Membership Header Data

        Dim vSQL As New StringBuilder
        vSQL.Append("INSERT INTO membership_header (contact_number,first_joined,first_membership_type,first_membership_source,first_number_of_members,first_membership_duration,")
        vSQL.Append("last_joined,last_membership_type,last_membership_source,last_number_of_members,last_membership_duration)")

        vSQL.Append(" SELECT fm.contact_number, fm.joined first_joined, fm.membership_type first_membership_type, fm.source first_membership_source,")
        vSQL.Append(pConn.DBIsNull("first_associate_count + FLOOR(fm.members_per_order /2) + 1", "FLOOR(fm.members_per_order /2) + 1"))
        vSQL.Append(" first_number_of_members,")
        vSQL.Append(" FLOOR(")
        vSQL.Append(pConn.DBIsNull(pConn.DBMonthDiff("fm.joined", "fm.cancelled_on"), pConn.DBMonthDiff("fm.joined", pConn.DBDate)))
        vSQL.Append(") first_membership_duration,")

        'All the 'last' items have to use ifnull as there may not be a 'last' single membership
        vSQL.Append(pConn.DBIsNull("lm.joined", "fm.joined"))
        vSQL.Append(" last_joined,")
        vSQL.Append(pConn.DBIsNull("lm.membership_type", "fm.membership_type"))
        vSQL.Append(" last_membership_type,")
        vSQL.Append(pConn.DBIsNull("lm.source", "fm.source"))
        vSQL.Append(" last_membership_source,")
        vSQL.Append(pConn.DBIsNull(pConn.DBIsNull("last_associate_count + FLOOR(lm.members_per_order /2) + 1", _
                                                  "FLOOR(lm.members_per_order /2) + 1"), _
                                   pConn.DBIsNull("first_associate_count + FLOOR(fm.members_per_order /2) + 1", _
                                                  "FLOOR(fm.members_per_order /2) + 1")))
        vSQL.Append(" last_number_of_members, ")
        vSQL.Append("FLOOR(")
        vSQL.Append(pConn.DBIsNull(pConn.DBIsNull(pConn.DBMonthDiff("lm.joined", "lm.cancelled_on"), _
                                                  pConn.DBMonthDiff("lm.joined", pConn.DBDate)), _
                                   pConn.DBIsNull(pConn.DBMonthDiff("fm.joined", "fm.cancelled_on"), _
                                                  pConn.DBMonthDiff("fm.joined", pConn.DBDate))))
        vSQL.Append(") last_membership_duration")
        vSQL.Append(" FROM contact_header ch")

        'Join to members and membership_types to get member data for first membership
        vSQL.Append(" INNER JOIN (SELECT contact_number, joined, cancelled_on, ms1.membership_type, source, order_number, membership_number, associate_membership_type, members_per_order")
        vSQL.Append(" FROM members ms1, membership_types mt1 WHERE ms1.membership_type = mt1.membership_type) fm ON fm.contact_number = ch.contact_number")

        'Find earliest member record by getting the lowest membership number for the lowest join date (may be multiple records for same join date)
        vSQL.Append(" INNER JOIN (SELECT contact_number, min(membership_number) membership_number")
        vSQL.Append(" FROM members mf1 WHERE joined = (SELECT min(joined) FROM members mf2 WHERE mf2.contact_number = mf1.contact_number)")
        vSQL.Append(" GROUP BY contact_number ) mino ON fm.contact_number = mino.contact_number AND fm.membership_number = mino.membership_number")

        'Get number of associates for each order whose membership type has an associated type
        'May not be any so outer join
        vSQL.Append(" LEFT OUTER JOIN (SELECT order_number, fam.membership_type, count(*) first_associate_count")
        vSQL.Append(" FROM members fam WHERE fam.membership_type IN (SELECT DISTINCT associate_membership_type FROM membership_types)")
        vSQL.Append(" GROUP BY order_number, fam.membership_type ) famc ON fm.order_number = famc.order_number AND fm.associate_membership_type = famc.membership_type")

        'Find the last member record for a single membership by getting the one with the highest join date and the highest membership number
        'May not be any so outer join
        vSQL.Append(" LEFT OUTER JOIN (select contact_number, max(membership_number) membership_number")
        vSQL.Append(" FROM members ml1, membership_types mlt1 WHERE ml1.membership_type = mlt1.membership_type AND mlt1.single_membership = 'Y'")
        vSQL.Append(" AND joined = (SELECT max(joined) FROM members ml2, membership_types mlt2")
        vSQL.Append(" WHERE ml2.membership_type = mlt2.membership_type AND mlt2.single_membership = 'Y' AND ml2.contact_number = ml1.contact_number)")
        vSQL.Append(" GROUP BY contact_number) maxo ON ch.contact_number = maxo.contact_number")

        'Join to members and membership_types to get member data for the last single membership
        'May not be any so outer join
        vSQL.Append(" LEFT OUTER JOIN (SELECT contact_number, joined, cancelled_on, ms2.membership_type, source, order_number, membership_number, associate_membership_type, members_per_order")
        vSQL.Append(" FROM members ms2, membership_types mt2 WHERE ms2.membership_type = mt2.membership_type) lm ON maxo.contact_number = lm.contact_number AND maxo.membership_number = lm.membership_number")

        'Get number of associates for each order whose membership type has an associated type
        'May not be any so outer join
        vSQL.Append(" LEFT OUTER JOIN (SELECT order_number, lam.membership_type, count(*) last_associate_count")
        vSQL.Append(" FROM members lam WHERE lam.membership_type IN (SELECT DISTINCT associate_membership_type FROM membership_types)")
        vSQL.Append(" GROUP BY order_number,lam.membership_type ) lamc ON lm.order_number = lamc.order_number AND lm.associate_membership_type = lamc.membership_type")

        If pContactNumber > 0 Then
          vSQL.Append(" WHERE lm.contact_number = ")
          vSQL.Append(pContactNumber)
        End If

        vCount = pConn.ExecuteSQL(pConn.ProcessAnsiJoins(vSQL.ToString))

        If pContactNumber = 0 And vLastMembershipNo = 0 Then
          LogStatus(pJob, (ProjectText.String30337)) 'Creating Indexes for Membership Header
          vCDBIndexes.ReCreate(pConn)
          vCDBIndexes.CreateIfMissing(pConn, False, {"contact_number"})
        End If

        If pContactNumber = 0 And vMaxMembershipNo > 0 Then
          vUpdateFields.Clear()
          vUpdateFields.Add("last_membership_number", CDBField.FieldTypes.cftLong, vMaxMembershipNo)
          vWhereFields.Clear()
          vWhereFields.Add("department", CDBField.FieldTypes.cftCharacter, pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDepartment))
          pConn.UpdateRecords("marketing_controls", vUpdateFields, vWhereFields, False)
        End If
        ProcessMembership = vCount
      Catch vEx As Exception
        PreserveStackTrace(vEx)
        vCDBIndexes.ReCreate(pConn)
        vCDBIndexes.CreateIfMissing(pConn, False, {"contact_number"})
        Throw vEx
      End Try
    End Function

    Public Function ProcessPerformances(ByVal pEnv As CDBEnvironment, ByVal pConn As CDBConnection, ByRef pJob As JobSchedule, ByVal pContactNumber As Integer, Optional ByRef pPerformance As String = "", Optional ByRef pSS As Integer = 0, Optional ByRef pSelectionTable As String = "") As Integer
      Dim vPerformances(0) As Performance
      Dim vRow As Integer
      Dim vPList As String = ""
      Dim vIndex As Integer
      Dim vCountMailings As Boolean
      Dim vProcessJoints As Boolean
      Dim vRecords As Integer
      Dim vPerformRecords As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vInsertFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vCDBIndexes As New CDBIndexes
      Dim vTempTable As String = ""
      Dim vCEWhere As String
      Dim vCEWhereBelow As String
      Dim vCEWhereAbove As String
      Dim vCEWhereBetween As String
      Dim vCEWhereRolling As String
      Dim vCEWherePreceding As String
      Dim vCMWhere As String
      Dim vPFields As String
      Dim vTempCreated As Boolean
      Dim vExpenditureGroups As StringList = Nothing

      Try
        LogStatus(pJob, (ProjectText.String30315)) 'Reading Performance Information

        Dim vUseHints As Boolean = pEnv.GetConfigOption("oracle_hints_mdg_performances")

        If pPerformance.Length > 0 Then
          vWhereFields.Add("performance", CDBField.FieldTypes.cftCharacter, pPerformance)
        Else
          vWhereFields.Add("automatic", CDBField.FieldTypes.cftCharacter, "Y")
        End If
        vRecordSet = pConn.GetRecordSet("SELECT * FROM performances WHERE " & pConn.WhereClause(vWhereFields) & " ORDER BY performance")
        While vRecordSet.Fetch() = True
          ReDim Preserve vPerformances(vRow)
          vPerformances(vRow) = New Performance
          With vPerformances(vRow)
            .Performance = vRecordSet.Fields("performance").Value
            If vRow > 0 Then vPList = vPList & ","
            vPList = vPList & "'" & .Performance & "'"
            .LowerLevel = vRecordSet.Fields("lower_level").DoubleValue
            .UpperLevel = vRecordSet.Fields("upper_level").DoubleValue
            .IncludeBefore = vRecordSet.Fields("include_payments_before").Value
            .IncludeAfter = vRecordSet.Fields("include_payments_after").Value
            .ExpenditureGroup = vRecordSet.Fields("expenditure_group").Value
            .RollingBoundary = vRecordSet.Fields("rolling_boundary").IntegerValue
            .ProcessMailings = vRecordSet.Fields("process_mailings").Bool
            If .ProcessMailings Then vCountMailings = True
            .MailingsBefore = vRecordSet.Fields("include_mailings_before").Value
            .MailingsAfter = vRecordSet.Fields("include_mailings_after").Value
            .IncludeJoint = vRecordSet.Fields("include_joint").Bool
            If .IncludeJoint Then vProcessJoints = True
            .AllPayMethods = vRecordSet.Fields("payment_method_list").Value
            .PaymentMethods = Split(.AllPayMethods, ",")
            'Since the list of payment methods may contain spaces around the commas and the list has been split based on the commas, make sure that each element in the array doesn' contain any spaces
            For vIndex = 0 To UBound(.PaymentMethods)
              .PaymentMethods(vIndex) = Trim(.PaymentMethods(vIndex))
            Next
            .RollingDateTo = CDate(TodaysDate())
            If IsDate(.IncludeBefore) And IsDate(.IncludeAfter) Then
              If CDate(.IncludeBefore) > CDate(.IncludeAfter) Then .RollingDateTo = CDate(.IncludeBefore)
            ElseIf IsDate(.IncludeBefore) And Not IsDate(.IncludeAfter) Then
              .RollingDateTo = CDate(.IncludeBefore)
            End If
            .RollingDateFrom = DateAdd(Microsoft.VisualBasic.DateInterval.Month, -.RollingBoundary, .RollingDateTo)
            .RollingDateFrom = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, .RollingDateFrom)

            .PrevRollingDateTo = DateAdd(Microsoft.VisualBasic.DateInterval.Month, -.RollingBoundary, .RollingDateTo)
            .PrevRollingDateFrom = DateAdd(Microsoft.VisualBasic.DateInterval.Month, -.RollingBoundary, .PrevRollingDateTo)
            .PrevRollingDateFrom = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, .PrevRollingDateFrom)
            vRow = vRow + 1
          End With
        End While
        vRecordSet.CloseRecordSet()

        If vPList = "" Then
          If pPerformance.Length > 0 Then
            RaiseError(DataAccessErrors.daeNoPerformance, pPerformance)
          Else
            RaiseError(DataAccessErrors.daeNoAutoPerformances)
          End If
        End If

        vWhereFields.Clear()
        If pContactNumber > 0 Then
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
        Else
          LogStatus(pJob, (ProjectText.String30316)) 'Dropping Indexes for Contact Performances
          vCDBIndexes.Init(pConn, "contact_performances")
          vCDBIndexes.DropAll(pConn)
        End If

        If InStr(vPList, ",") > 0 Then
          vWhereFields.Add("performance", CDBField.FieldTypes.cftCharacter, vPList, CDBField.FieldWhereOperators.fwoIn)
        Else
          vWhereFields.Add("performance", CDBField.FieldTypes.cftCharacter, Replace(vPList, "'", ""))
        End If
        LogStatus(pJob, String.Format(ProjectText.String30317, vPList)) 'Deleting Performance Records for: %s
        pConn.DeleteRecords("contact_performances", vWhereFields, False)

        '-----------------------------------------------------------------------------------
        'New code here to try to create all performance data in one go
        '-----------------------------------------------------------------------------------
        vPFields = "contact_number, performance, number_of_payments, value_of_payments,"
        vPFields = vPFields & "number_of_mailings,no_response,average_value,average_per_mailing,response_rate,number_above,"
        vPFields = vPFields & "value_above,number_below,value_below,number_between,value_between,rolling_value,preceding_rolling_value,"
        vPFields = vPFields & "first_payment,first_payment_date,last_payment,last_payment_date,maximum_payment,maximum_payment_date"

        For vRow = 0 To UBound(vPerformances)
          With vPerformances(vRow)
            vCMWhere = "" 'Restrictions for contact mailings
            If IsDate(.MailingsAfter) Then
              If IsDate(.MailingsBefore) Then
                vCMWhere = "mailing_date" & pConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, .MailingsAfter) & " AND mailing_date" & pConn.SQLLiteral("<", CDBField.FieldTypes.cftDate, .MailingsBefore)
              Else
                vCMWhere = "mailing_date" & pConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, .MailingsAfter)
              End If
            ElseIf IsDate(.MailingsBefore) Then
              vCMWhere = "mailing_date" & pConn.SQLLiteral("<", CDBField.FieldTypes.cftDate, .MailingsBefore)
            End If
            'Exclude the rows with the wrong expenditure groug
            If .ExpenditureGroup.Length > 0 Then
              vExpenditureGroups = New StringList(.ExpenditureGroup)
              If Len(vCMWhere) > 0 Then vCMWhere = vCMWhere & " AND "
              vCMWhere = vCMWhere & "expenditure_group IN(" & vExpenditureGroups.InList & ")"
            End If
            If Len(vCMWhere) > 0 Then vCMWhere = " AND " & vCMWhere
            vCMWhere = " WHERE marketing = 'Y'" & vCMWhere

            vCEWhere = "" 'Restrictions for expenditure
            If IsDate(.IncludeAfter) Then
              If IsDate(.IncludeBefore) Then
                vCEWhere = "transaction_date" & pConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, .IncludeAfter) & " AND transaction_date" & pConn.SQLLiteral("<", CDBField.FieldTypes.cftDate, .IncludeBefore)
              Else
                vCEWhere = "transaction_date" & pConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, .IncludeAfter)
              End If
            ElseIf IsDate(.IncludeBefore) Then
              vCEWhere = "transaction_date" & pConn.SQLLiteral("<", CDBField.FieldTypes.cftDate, .IncludeBefore)
            End If
            'Exclude the rows with the wrong expenditure groug
            If .ExpenditureGroup.Length > 0 Then
              If Len(vCEWhere) > 0 Then vCEWhere = vCEWhere & " AND "
              vCEWhere = vCEWhere & "expenditure_group IN(" & vExpenditureGroups.InList & ")"
            End If
            'Exclude the rows with the wrong payment methods
            If .AllPayMethods.Length > 0 Then
              If Len(vCEWhere) > 0 Then vCEWhere = vCEWhere & " AND "
              vCEWhere = vCEWhere & "payment_method IN ("
              For vIndex = 0 To UBound(.PaymentMethods)
                vCEWhere = vCEWhere & "'" & .PaymentMethods(vIndex) & "'"
                If vIndex < UBound(.PaymentMethods) Then vCEWhere = vCEWhere & ","
              Next
              vCEWhere = vCEWhere & ")"
            End If

            If Len(vCEWhere) > 0 Then
              vCEWhere = " WHERE " & vCEWhere
              vCEWhereAbove = vCEWhere & " AND "
              vCEWhereBelow = vCEWhere & " AND "
              vCEWhereBetween = vCEWhere & " AND "
              vCEWhereRolling = vCEWhere & " AND "
              vCEWherePreceding = vCEWhere & " AND "
            Else
              vCEWhereAbove = " WHERE "
              vCEWhereBelow = " WHERE "
              vCEWhereBetween = " WHERE "
              vCEWhereRolling = " WHERE "
              vCEWherePreceding = " WHERE "
            End If
            vCEWhereAbove = vCEWhereAbove & "amount >= " & .UpperLevel
            vCEWhereBelow = vCEWhereBelow & "amount <= " & .LowerLevel
            vCEWhereBetween = vCEWhereBetween & "amount > " & .LowerLevel & " AND amount < " & .UpperLevel
            vCEWhereRolling = vCEWhereRolling & "transaction_date" & pConn.SQLLiteral(">=", .RollingDateFrom) & " AND transaction_date" & pConn.SQLLiteral("<=", .RollingDateTo)
            vCEWherePreceding = vCEWherePreceding & "transaction_date" & pConn.SQLLiteral(">=", .PrevRollingDateFrom) & " AND transaction_date" & pConn.SQLLiteral("<=", .PrevRollingDateTo)

            Dim vSQL As New StringBuilder
            vSQL.Append("INSERT INTO contact_performances (")
            vSQL.Append(vPFields)
            vSQL.Append(")")
            vSQL.Append(" SELECT ")
            vSQL.Append(pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtUseHashOracle9, "ch,cmrc", vUseHints))
            vSQL.Append(" ch.contact_number,'")
            vSQL.Append(.Performance)
            vSQL.Append("',")
            vSQL.Append(pConn.DBIsNull("number_of_payments", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("value_of_payments", "0"))
            vSQL.Append(",")
            If .ProcessMailings Then
              vSQL.Append(pConn.DBIsNull("number_of_mailings", "0"))
              vSQL.Append(",")
              vSQL.Append(pConn.DBIsNull("number_of_mailings - " & pConn.DBIsNull("got_response_count", "0"), "0"))
              vSQL.Append(",")
            Else
              vSQL.Append("0,0,")
            End If
            vSQL.Append(pConn.DBIsNull("value_of_payments / number_of_payments", "0"))
            vSQL.Append(",")
            If .ProcessMailings Then
              vSQL.Append(pConn.DBIsNull("value_of_payments / number_of_mailings", "0"))
              vSQL.Append(",")
              vSQL.Append(pConn.DBIsNull(pConn.DBToNumber("number_of_payments") & " / number_of_mailings", "0"))
              vSQL.Append(",")
            Else
              vSQL.Append("0,0,")
            End If
            vSQL.Append(pConn.DBIsNull("number_above", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("value_above", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("number_below", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("value_below", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("number_between", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("value_between", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("rolling_value", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("preceding_rolling_value", "0"))
            vSQL.Append(",")
            vSQL.Append(pConn.DBIsNull("first_payment", "0"))
            vSQL.Append(",first_payment_date,")
            vSQL.Append(pConn.DBIsNull("last_payment", "0"))
            vSQL.Append(",last_payment_date,")
            vSQL.Append(pConn.DBIsNull("maximum_payment", "0"))
            vSQL.Append(",maximum_payment_date")
            If pSS > 0 Then
              vSQL.Append(" FROM " & pSelectionTable & " sc INNER JOIN contact_header ch ON sc.contact_number = ch.contact_number")
            Else
              vSQL.Append(" FROM contact_header ch")
            End If

            vSQL.Append(" LEFT OUTER JOIN (SELECT te.contact_number, number_of_payments, value_of_payments,")
            vSQL.Append(" first_payment_date, last_payment_date, maximum_payment, first_payment, last_payment FROM ")

            vSQL.Append(" (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "ce", vUseHints) & " contact_number, count(*) AS number_of_payments, SUM(amount) AS value_of_payments,")
            vSQL.Append(" MIN(transaction_date) AS first_payment_date, MAX(transaction_date) AS last_payment_date,MAX(Amount) AS maximum_payment")
            vSQL.Append(" FROM contact_expenditure ce " & vCEWhere)
            vSQL.Append(" GROUP BY ce.contact_number) te")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "contact_expenditure", vUseHints) & " SUM(amount) AS first_payment, transaction_date, contact_number")
            vSQL.Append(" FROM contact_expenditure " & vCEWhere)
            vSQL.Append(" GROUP BY contact_number, transaction_date) cefd")
            vSQL.Append(" ON te.contact_number = cefd.contact_number AND te.first_payment_date = cefd.transaction_date")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "contact_expenditure", vUseHints) & " SUM(amount) AS last_payment, transaction_date, contact_number")
            vSQL.Append(" FROM contact_expenditure " & vCEWhere)
            vSQL.Append(" GROUP BY contact_number, transaction_date) celd")
            vSQL.Append(" ON te.contact_number = celd.contact_number AND te.last_payment_date = celd.transaction_date ) tefl")
            vSQL.Append(" ON ch.contact_number = tefl.contact_number")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "cempd", vUseHints) & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtNoMergeOracle9, "ma", vUseHints) & " cempd.contact_number, MAX( transaction_date) AS maximum_payment_date")
            vSQL.Append(" FROM contact_expenditure cempd INNER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "cema", vUseHints) & " cema.contact_number, MAX(amount) AS max_amount")
            vSQL.Append(" FROM contact_expenditure cema " & vCEWhere)
            vSQL.Append(" GROUP BY cema.contact_number) ma ON cempd.contact_number = ma.contact_number AND cempd.amount = ma.max_amount" & vCEWhere)
            vSQL.Append(" GROUP BY cempd.contact_number) mtd ON tefl.contact_number = mtd.contact_number")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "contact_expenditure", vUseHints) & " contact_number, SUM(amount) AS value_below, count(*) as number_below")
            vSQL.Append(" FROM contact_expenditure" & vCEWhereBelow)
            vSQL.Append(" GROUP BY contact_number) nvb  ON tefl.contact_number = nvb.contact_number")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "contact_expenditure", vUseHints) & " contact_number, SUM(amount) AS value_above, count(*) as number_above")
            vSQL.Append(" FROM contact_expenditure" & vCEWhereAbove)
            vSQL.Append(" GROUP BY contact_number) nva ON tefl.contact_number = nva.contact_number")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "contact_expenditure", vUseHints) & " contact_number, SUM(amount) AS value_between, count(*) as number_between")
            vSQL.Append(" FROM contact_expenditure" & vCEWhereBetween)
            vSQL.Append(" GROUP BY contact_number) nvm ON tefl.contact_number = nvm.contact_number")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "contact_expenditure", vUseHints) & " contact_number, SUM(amount) AS rolling_value")
            vSQL.Append(" FROM contact_expenditure" & vCEWhereRolling)
            vSQL.Append(" GROUP BY contact_number) erv ON tefl.contact_number = erv.contact_number")

            vSQL.Append(" LEFT OUTER JOIN (SELECT" & pConn.DBHint(CDBConnection.DatabaseHintTypes.dhtFullTableScanOracle8, "contact_expenditure", vUseHints) & " contact_number, SUM(amount) AS preceding_rolling_value")
            vSQL.Append(" FROM contact_expenditure" & vCEWherePreceding)
            vSQL.Append(" GROUP BY contact_number) prv ON tefl.contact_number = prv.contact_number")

            If .ProcessMailings Then
              vSQL.Append(" LEFT OUTER JOIN (SELECT cm.contact_number,count(*) AS number_of_mailings FROM contact_mailings cm")
              vSQL.Append(" INNER JOIN mailing_history mh ON cm.mailing_number = mh.mailing_number")
              vSQL.Append(" INNER JOIN mailings m ON mh.mailing = m.mailing")
              vSQL.Append(" INNER JOIN segments s ON cm.mailing_number = s.mailing_number")
              vSQL.Append(" INNER JOIN appeals a ON s.campaign = a.campaign AND s.appeal = a.appeal")
              vSQL.Append(vCMWhere & " GROUP BY contact_number) cmc ON ch.contact_number = cmc.contact_number")

              vSQL.Append(" LEFT OUTER JOIN (SELECT cm.contact_number,count(*) AS got_response_count FROM contact_mailings cm")
              vSQL.Append(" INNER JOIN mailing_history mh ON cm.mailing_number = mh.mailing_number")
              vSQL.Append(" INNER JOIN mailings m ON mh.mailing = m.mailing")
              vSQL.Append(" INNER JOIN segments s ON cm.mailing_number = s.mailing_number")
              vSQL.Append(" INNER JOIN appeals a ON s.campaign = a.campaign AND s.appeal = a.appeal")
              vSQL.Append(" INNER JOIN (SELECT cemr.contact_number, MAX(transaction_date) AS last_date")
              vSQL.Append(" FROM contact_expenditure cemr " & vCEWhere)
              vSQL.Append(" GROUP BY cemr.contact_number) mr ON cm.contact_number = mr.contact_number AND mh.mailing_date <= mr.last_date") 'BR20134 Handle response on the same date as the mailing < to <=
              vSQL.Append(vCMWhere & " GROUP BY cm.contact_number) cmrc ON ch.contact_number = cmrc.contact_number")
            End If

            If pContactNumber > 0 Then
              vSQL.Append(" WHERE ch.contact_number = " & pContactNumber)
            Else
              If pSS > 0 Then vSQL.Append(" WHERE sc.selection_set = " & pSS & " AND revision = 1 ")
            End If
            LogStatus(pJob, String.Format(ProjectText.String30338, .Performance)) '(ProjectText.String30319)     'Selecting Expenditure Data for Performances    'Selecting Expenditure Data for Performance %s
            pJob.RecordType = (ProjectText.String30339) 'Performance Records
            vPerformRecords = vPerformRecords + pConn.ExecuteSQL(pConn.ProcessAnsiJoins(vSQL.ToString))
            pJob.RecordsProcessed = vPerformRecords
          End With
        Next

        If pContactNumber <= 0 Then
          LogStatus(pJob, (ProjectText.String30322)) 'Creating Indexes for Contact Performances
          vCDBIndexes.ReCreate(pConn)
          vCDBIndexes.CreateIfMissing(pConn, True, {"contact_number", "performance"})
        End If

        For vRow = 0 To UBound(vPerformances)
          With vPerformances(vRow)
            If .IncludeJoint Then
              LogStatus(pJob, (ProjectText.String30323)) 'Processing Contact Performances for Joint Contacts
              If pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink) = "" Or pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink) = "" Then
                RaiseError(DataAccessErrors.daeJointCodesUndefined) 'Joint Relationship Codes are not defined
              End If

              vTempTable = "rpt_temp_contact_performances"
              'Drop temporary table if already exists
              If pConn.TableExists(vTempTable) Then pConn.DropTable(vTempTable)
              'Create new temporary table
              vInsertFields.Clear()
              vInsertFields.Add("contact_number", CDBField.FieldTypes.cftLong)
              vInsertFields.Add("performance", CDBField.FieldTypes.cftCharacter, "6")
              vInsertFields.Add("number_of_payments", CDBField.FieldTypes.cftLong)
              vInsertFields.Add("value_of_payments", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("number_of_mailings", CDBField.FieldTypes.cftLong)
              vInsertFields.Add("no_response", CDBField.FieldTypes.cftLong)
              vInsertFields.Add("average_value", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("average_per_mailing", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("response_rate", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("number_above", CDBField.FieldTypes.cftLong)
              vInsertFields.Add("value_above", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("number_below", CDBField.FieldTypes.cftLong)
              vInsertFields.Add("value_below", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("number_between", CDBField.FieldTypes.cftLong)
              vInsertFields.Add("value_between", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("rolling_value", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("preceding_rolling_value", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("first_payment", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("first_payment_date", CDBField.FieldTypes.cftDate)
              vInsertFields.Add("last_payment", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("last_payment_date", CDBField.FieldTypes.cftDate)
              vInsertFields.Add("maximum_payment", CDBField.FieldTypes.cftNumeric, "11")
              vInsertFields.Add("maximum_payment_date", CDBField.FieldTypes.cftDate)
              pConn.CreateTableFromFields(vTempTable, vInsertFields)
              vTempCreated = True

              'Insert the contacts who have a joint contact also in the performances table
              Dim vSQL As New StringBuilder
              vSQL.Append("INSERT INTO ")
              vSQL.Append(vTempTable)
              vSQL.Append(" (")
              vSQL.Append(vPFields)
              vSQL.Append(") ")
              vSQL.Append("SELECT DISTINCT cp.contact_number,cp.performance,cp.number_of_payments,cp.value_of_payments,")
              vSQL.Append("cp.number_of_mailings,cp.no_response,cp.average_value,cp.average_per_mailing,cp.response_rate,cp.number_above,")
              vSQL.Append("cp.value_above,cp.number_below,cp.value_below,cp.number_between,cp.value_between,cp.rolling_value,cp.preceding_rolling_value,")
              vSQL.Append("cp.first_payment , cp.first_payment_date, cp.last_payment, cp.last_payment_date, cp.maximum_payment, cp.maximum_payment_date")
              vSQL.Append(" FROM contact_performances cp, contact_links cl, contact_performances jp WHERE")
              If pContactNumber > 0 Then
                vSQL.Append(" cp.contact_number = ")
                vSQL.Append(pContactNumber)
                vSQL.Append(" AND ")
              End If
              vSQL.Append(" cp.performance = '")
              vSQL.Append(.Performance)
              vSQL.Append("' AND cp.contact_number = cl.contact_number_2 AND cl.relationship IN ('")
              vSQL.Append(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink) & "', '")
              vSQL.Append(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink) & "')")
              vSQL.Append(" AND jp.contact_number = cl.contact_number_1 AND jp.performance = cp.performance")
              vRecords = pConn.ExecuteSQL(vSQL.ToString)
              Debug.Print(vRecords & " Individuals moved to Temporary Table")

              'Now insert the joint contacts
              vSQL = New StringBuilder
              vSQL.Append("INSERT INTO " & vTempTable & " (" & vPFields & ") ")
              vSQL.Append("SELECT cp.contact_number,cp.performance,jp.number_of_payments,jp.value_of_payments,")
              vSQL.Append("jp.number_of_mailings,jp.no_response,jp.average_value,jp.average_per_mailing,jp.response_rate,jp.number_above,")
              vSQL.Append("jp.value_above,jp.number_below,jp.value_below,jp.number_between,jp.value_between,jp.rolling_value,jp.preceding_rolling_value,")
              vSQL.Append("jp.first_payment , jp.first_payment_date, jp.last_payment, jp.last_payment_date, jp.maximum_payment, jp.maximum_payment_date")
              vSQL.Append(" FROM contact_performances cp, contact_links cl, contact_performances jp WHERE")
              If pContactNumber > 0 Then vSQL.Append(" cp.contact_number = " & pContactNumber & " AND ")
              vSQL.Append(" cp.performance = '" & .Performance & "' AND cp.contact_number = cl.contact_number_2 AND cl.relationship IN ('")
              vSQL.Append(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink) & "', '")
              vSQL.Append(pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink) & "')")
              vSQL.Append(" AND jp.contact_number = cl.contact_number_1 AND jp.performance = cp.performance")
              vRecords = pConn.ExecuteSQL(vSQL.ToString)
              Debug.Print(vRecords & " Joints moved to Temporary Table")

              'Create an index on the temporary table
              pConn.CreateIndex(False, vTempTable, {"contact_number"})

              'Delete the individuals from the original table
              vSQL = New StringBuilder
              vSQL.Append("DELETE FROM contact_performances WHERE contact_number IN (SELECT DISTINCT contact_number FROM " & vTempTable & ") AND performance = '" & .Performance & "'")
              vRecords = pConn.ExecuteSQL(vSQL.ToString)
              Debug.Print(vRecords & " Performance Records Deleted")

              'Now select the aggregate data
              vSQL = New StringBuilder
              vSQL.Append("INSERT INTO contact_performances (" & vPFields & ") ")
              vSQL.Append("SELECT te.contact_number, '" & .Performance & "', number_of_payments, value_of_payments,")
              vSQL.Append("number_of_mailings, no_response, 0, 0, 0,")
              vSQL.Append("number_above, value_above, number_below, value_below, number_between, value_between,")
              vSQL.Append("rolling_value, preceding_rolling_value,")
              vSQL.Append(pConn.DBIsNull("first_payment", "0"))
              vSQL.Append(", te.first_payment_date,")
              vSQL.Append(pConn.DBIsNull("last_payment", "0"))
              vSQL.Append(", te.last_payment_date, ")
              vSQL.Append(pConn.DBIsNull("maximum_payment", "0"))
              vSQL.Append(", maximum_payment_date")
              vSQL.Append(" FROM (SELECT contact_number, '" & .Performance & "' as performance, SUM(number_of_payments) AS number_of_payments, SUM(value_of_payments) AS value_of_payments,")
              vSQL.Append(" SUM(number_of_mailings) AS number_of_mailings, SUM(no_response) AS no_response,")
              vSQL.Append(" SUM(number_above) AS number_above, SUM(value_above) AS value_above,")
              vSQL.Append(" SUM(number_below) AS number_below, SUM(value_below) AS value_below,")
              vSQL.Append(" SUM(number_between) AS number_between, SUM(value_between) AS value_between,")
              vSQL.Append(" SUM(rolling_value) AS rolling_value, SUM(preceding_rolling_value) AS preceding_rolling_value,")
              vSQL.Append(" MIN(first_payment_date) AS first_payment_date, MAX(last_payment_date) AS last_payment_date, MAX(maximum_payment) AS maximum_payment")
              vSQL.Append(" FROM " & vTempTable & " GROUP BY contact_number ) te")

              vSQL.Append(" LEFT OUTER JOIN (SELECT tpmpd.contact_number, MAX(maximum_payment_date) AS maximum_payment_date FROM " & vTempTable & " tpmpd")
              vSQL.Append(" INNER JOIN (SELECT tpma.contact_number, MAX(maximum_payment) AS max_amount FROM " & vTempTable & " tpma GROUP BY tpma.contact_number) ma")
              vSQL.Append(" ON tpmpd.contact_number = ma.contact_number AND tpmpd.maximum_payment = ma.max_amount")
              vSQL.Append(" GROUP BY tpmpd.contact_number) mtd ON te.contact_number = mtd.contact_number")

              vSQL.Append(" LEFT OUTER JOIN (SELECT SUM(first_payment) AS first_payment, first_payment_date, contact_number")
              vSQL.Append(" FROM " & vTempTable & " GROUP BY contact_number, first_payment_date) tpfd ON te.contact_number = tpfd.contact_number AND te.first_payment_date = tpfd.first_payment_date")

              vSQL.Append(" LEFT OUTER JOIN (SELECT SUM(last_payment) AS last_payment, last_payment_date, contact_number")
              vSQL.Append(" FROM " & vTempTable & " GROUP BY contact_number, last_payment_date) tpld ON te.contact_number = tpld.contact_number AND te.last_payment_date = tpld.last_payment_date")

              vRecords = pConn.ExecuteSQL(pConn.ProcessAnsiJoins(vSQL.ToString))
              Debug.Print(vRecords & " Aggregate records restored")

              'Update the average value
              vSQL = New StringBuilder
              vSQL.Append("UPDATE contact_performances SET average_value = value_of_payments / number_of_payments")
              vSQL.Append(" WHERE performance = '" & .Performance & "' AND contact_number IN ( SELECT DISTINCT contact_number FROM " & vTempTable & ") AND number_of_payments > 0")
              pConn.ExecuteSQL(vSQL.ToString)

              If .ProcessMailings Then
                'If processing mailings then update the average per mailing and the response rate
                vSQL = New StringBuilder
                vSQL.Append("UPDATE contact_performances SET average_per_mailing = value_of_payments / number_of_mailings, response_rate = number_of_payments / number_of_mailings")
                vSQL.Append(" WHERE performance = '" & .Performance & "' AND contact_number IN ( SELECT DISTINCT contact_number FROM " & vTempTable & ") AND number_of_mailings > 0")
                pConn.ExecuteSQL(vSQL.ToString)
              End If
            End If
          End With
        Next
        If vTempCreated Then pConn.DropTable(vTempTable)
        vTempCreated = False

        'Make sure response rate is 9.9 or less
        vWhereFields = New CDBFields
        If InStr(vPList, ",") > 0 Then
          vWhereFields.Add("performance", CDBField.FieldTypes.cftCharacter, vPList, CDBField.FieldWhereOperators.fwoIn)
        Else
          vWhereFields.Add("performance", CDBField.FieldTypes.cftCharacter, Replace(vPList, "'", ""))
        End If
        vWhereFields.Add("response_rate", CDBField.FieldTypes.cftNumeric, "9.9", CDBField.FieldWhereOperators.fwoGreaterThan)
        vUpdateFields = New CDBFields
        vUpdateFields.Add("response_rate", CDBField.FieldTypes.cftNumeric, "9.9")
        pConn.UpdateRecords("contact_performances", vUpdateFields, vWhereFields, False)
        ProcessPerformances = vPerformRecords

      Catch vEx As Exception
        PreserveStackTrace(vEx)
        vCDBIndexes.ReCreate(pConn)
        vCDBIndexes.CreateIfMissing(pConn, True, {"contact_number", "performance"})
        If vTempCreated Then pConn.DropTable(vTempTable)
        Throw vEx
      End Try
    End Function

    Private Function LimitValue(ByRef pValue As Double) As Double
      If pValue > 9999999.99 Then
        LimitValue = 9999999.99
      Else
        LimitValue = pValue
      End If
    End Function

    Public Function ProcessScores(ByVal pEnv As CDBEnvironment, ByVal pConn As CDBConnection, ByRef pJob As JobSchedule, ByVal pContactNumber As Integer, Optional ByRef pScore As String = "", Optional ByRef pSS As Integer = 0, Optional ByRef pSelectionTable As String = "") As Integer
      Dim vScoringRows() As ScoringRow
      Dim vRow As Integer
      Dim vRS As CDBRecordSet
      Dim vRS2 As CDBRecordSet
      Dim vValue As String
      Dim vLastScore As String = ""
      Dim vWrite As Boolean
      Dim vContactNumber As Integer
      Dim vFound As Boolean
      Dim vPoints As Double
      Dim vCHRS As CDBRecordSet = Nothing
      Dim vCHEOF As Boolean
      Dim vWhere As String
      Dim vAllContacts As Boolean
      Dim vDone As Boolean
      Dim vMain As String
      Dim vSub As String = ""
      Dim vFrom As String = ""
      Dim vTo As String = ""
      Dim vAttrs As String
      Dim vContactWhere As String = ""
      Dim vOrgWhere As String = ""
      Dim vFromSQL As String
      Dim vRecords As Integer
      Dim vScores As String = ""
      Dim vDate As String
      Dim vTemp As String
      Dim vContactTable As String
      Dim vInsertFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vSQL As String
      Dim vCDBIndexes As New CDBIndexes
      Dim vMaintAttr As String
      Dim vSubsidiaryAttr As String
      Dim vFromAttr As String
      Dim vToAttr As String

      Try
        ReDim vScoringRows(0)
        If pContactNumber > 0 Then
          ProcessHeader(pEnv, pConn, pJob, pContactNumber)
          ProcessExpenditure(pEnv, pConn, pJob, pContactNumber)
          ProcessMembership(pEnv, pConn, pJob, pContactNumber)
          ProcessPerformances(pEnv, pConn, pJob, pContactNumber)
        End If

        LogStatus(pJob, (ProjectText.String30326)) 'Reading Score Information
        If pContactNumber > 0 Then
          vAllContacts = True
          vContactWhere = "contact_number = " & pContactNumber
          vOrgWhere = "organisation_number = " & pContactNumber
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
        End If

        'OLD code used to clear all scoring exceptions at this point

        vInsertFields.Add("contact_number", CDBField.FieldTypes.cftLong)
        vInsertFields.Add("score")
        vInsertFields.Add("points", CDBField.FieldTypes.cftNumeric)

        'Get the details for all the scores - all the automatic scores present
        If pScore.Length > 0 Then
          vWhere = "s.score = '" & pScore & "'"
        Else
          vWhere = "automatic = 'Y'"
        End If
        vRS = pConn.GetRecordSet("SELECT s.score,sd.search_area,i_e,sd.c_o,main_attribute,subsidiary_attribute,from_attribute,to_attribute,main_value,subsidiary_value,period,points,multiplier,table_name,main_data_type,subsidiary_data_type,special FROM scores s, scoring_details sd, selection_control sc WHERE " & vWhere & " AND s.score = sd.score AND sd.search_area = sc.search_area AND sd.c_o = sc.c_o AND sc.application_name = 'SM' ORDER BY s.score")
        While vRS.Fetch() = True
          ReDim Preserve vScoringRows(vRow)
          vScoringRows(vRow) = New ScoringRow
          With vScoringRows(vRow)
            .Score = vRS.Fields("score").Value
            .SearchAreaName = vRS.Fields("search_area").Value
            .IorE = vRS.Fields("i_e").Value
            .CorO = vRS.Fields("c_o").Value
            If .CorO = "O" Then
              .ContactAttr = "organisation_number"
            Else
              .ContactAttr = "contact_number"
            End If
            .MainAttribute = vRS.Fields("main_attribute").Value
            .SubsidiaryAttribute = vRS.Fields("subsidiary_attribute").Value
            .FromAttribute = vRS.Fields("from_attribute").Value
            .ToAttribute = vRS.Fields("to_attribute").Value
            If ParseValue((vRS.Fields("main_value").Value), .MainValue, .MainValueTo) Then RaiseError(DataAccessErrors.daeScoreInvalidSearchArea, .Score, .Score, "Main Value")
            If ParseValue((vRS.Fields("subsidiary_value").Value), .SubsidiaryValue, .SubsidiaryValueTo) Then RaiseError(DataAccessErrors.daeScoreInvalidSearchArea, .Score, .Score, "Subsidiary Value")
            If ParseValue((vRS.Fields("period").Value), .PeriodValue, .PeriodValueTo) Then RaiseError(DataAccessErrors.daeScoreInvalidSearchArea, .Score, .Score, "Period")
            .MainDataType = Trim(vRS.Fields("main_data_type").Value)
            .SubsidiaryDataType = Trim(vRS.Fields("subsidiary_data_type").Value)
            vValue = vRS.Fields("multiplier").Value
            If vValue.Length > 0 Then
              .Multiplier = Val(vValue)
              .ValueRequired = True
            Else
              .Multiplier = 0
              .ValueRequired = False
            End If
            .Points = vRS.Fields("points").DoubleValue
            .TableName = vRS.Fields("table_name").Value
            If vRS.Fields("special").Bool Then 'NoTranslate
              vRS2 = pConn.GetRecordSet("SELECT * FROM selection_control_details WHERE search_area = '" & .SearchAreaName & "' AND application_name = 'SM' AND c_o = '" & .CorO & "' ORDER BY sequence_number DESC")
              While vRS2.Fetch() = True
                vTemp = vRS2.Fields("table_2").Value
                If vTemp.Length > 0 Then
                  If .SpecialTables = "" Then .FirstSpecial = vTemp
                  .SpecialTables = .SpecialTables & vTemp & ", "
                  If Len(.SpecialLink) > 0 Then .SpecialLink = .SpecialLink & " AND "
                  .SpecialLink = .SpecialLink & vTemp & "." & vRS2.Fields("attribute_2").Value
                  .SpecialLink = .SpecialLink & " " & vRS2.Fields("join_condition").Value & " "
                  .SpecialLink = .SpecialLink & vRS2.Fields("table_1").Value & "." & vRS2.Fields("attribute_1").Value
                Else
                  .SpecialLink = .SpecialLink & vRS2.Fields("table_1").Value & "." & vRS2.Fields("attribute_1").Value
                  .SpecialLink = .SpecialLink & " " & vRS2.Fields("join_condition").Value
                End If
              End While
              vRS2.CloseRecordSet()
            End If
            Select Case .TableName
              Case "contact_header"
                .ContactHeader = True
            End Select
            vRow = vRow + 1
          End With
        End While
        vRS.CloseRecordSet()

        'Remove existing scores
        For vRow = 0 To UBound(vScoringRows)
          If vScoringRows(vRow) IsNot Nothing Then
            With vScoringRows(vRow)
              If .Score <> vLastScore Then
                If vScores.Length > 0 Then vScores = vScores & ","
                vScores = vScores & "'" & .Score & "'"
                vLastScore = .Score
              End If
            End With
          End If
        Next
        If InStr(vScores, ",") > 0 Then
          vWhereFields.Add("score", CDBField.FieldTypes.cftCharacter, vScores, CDBField.FieldWhereOperators.fwoIn)
        Else
          vWhereFields.Add("score", CDBField.FieldTypes.cftCharacter, Replace(vScores, "'", ""))
        End If
        vLastScore = ""

        If vScores = "" Then
          If pScore.Length > 0 Then
            RaiseError(DataAccessErrors.daeNoScore, pScore)
          Else
            RaiseError(DataAccessErrors.daeNoAutoScores)
          End If
        Else
          If pContactNumber = 0 Then
            LogStatus(pJob, (ProjectText.String30328)) 'Dropping Indexes for Contact Scores
            vCDBIndexes.Init(pConn, "contact_scores")
            vCDBIndexes.DropAll(pConn)
          End If
          LogStatus(pJob, String.Format(ProjectText.String30329, vScores)) 'Deleting Score Records for: %s
          pConn.DeleteRecords("contact_scores", vWhereFields, False)
        End If

        pJob.RecordType = (ProjectText.String30340) 'Scoring Records
        'Now we should set up cursors to select the records that will give us the required information
        'This may be for an individual contact, a selection set, or all contacts
        'For the present let us assume that each scoring row can be handled by a single select
        LogStatus(pJob, (ProjectText.String30330)) 'Selecting Contacts
        If vAllContacts Then
          If pContactNumber > 0 Then vWhere = "WHERE contact_number = " & pContactNumber
          vRS = pConn.GetRecordSet("SELECT contact_number FROM contacts " & vWhere & " ORDER BY contact_number", CDBConnection.RecordSetOptions.MultipleResultSets)
        Else
          If pSS > 0 Then
            vSQL = "SELECT ch.contact_number FROM " & pSelectionTable & " sc, contact_header ch WHERE sc.selection_set = " & pSS & " AND revision = 1 AND sc.contact_number = ch.contact_number ORDER BY sc.contact_number"
          Else
            vSQL = "SELECT * FROM contact_header ORDER BY contact_number"
          End If
          vRS = pConn.GetRecordSet(vSQL, CDBConnection.RecordSetOptions.MultipleResultSets)
          If pSS <= 0 Then vCHRS = vRS
        End If
        If vRS.Fetch() = True Then
          vContactNumber = CInt(vRS.Fields("contact_number").Value)
        Else
          If Not vAllContacts Then vCHEOF = True
          vDone = True
        End If

        For vRow = 0 To UBound(vScoringRows)
          With vScoringRows(vRow)
            LogStatus(pJob, String.Format(ProjectText.String30331, .Score)) 'Selecting Data For Score: %s
            If .ContactHeader = True Then
              If vCHRS Is Nothing Then
                If vContactWhere.Length > 0 Then vWhere = " WHERE " & vContactWhere
                vCHRS = pConn.GetRecordSet("SELECT * FROM contact_header " & vWhere & " ORDER BY contact_number")
                If vCHRS.Fetch() = False Then vCHEOF = True
              End If
              .RecordSet = vCHRS
            Else
              vWhere = ""
              vMaintAttr = .TableName & "." & .MainAttribute
              vSubsidiaryAttr = .TableName & "." & .SubsidiaryAttribute
              vFromAttr = .TableName & "." & .FromAttribute
              vToAttr = .TableName & "." & .ToAttribute
              If .MainValueTo.Length > 0 Then 'It's a range
                vWhere = vWhere & vMaintAttr & BetweenAttrValue(.MainDataType, .MainValue, .MainValueTo)
              Else
                vWhere = vWhere & vMaintAttr & InAttrValue(.MainDataType, .MainValue)
              End If
              If Len(.SubsidiaryValueTo) > 0 Then 'It's a range
                vWhere = vWhere & " AND " & vSubsidiaryAttr & BetweenAttrValue(.SubsidiaryDataType, .SubsidiaryValue, .SubsidiaryValueTo)
              ElseIf Len(.SubsidiaryValue) > 0 Then
                vWhere = vWhere & " AND " & vSubsidiaryAttr & InAttrValue(.SubsidiaryDataType, .SubsidiaryValue)
              End If
              If Len(.PeriodValueTo) > 0 Then 'It's a range
                If .ToAttribute.Length > 0 Then
                  vWhere = vWhere & " AND " & vFromAttr & pConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, .PeriodValue) & " AND " & vToAttr & pConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, .PeriodValueTo)
                Else
                  vWhere = vWhere & " AND " & vFromAttr & pConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, .PeriodValue) & " AND " & vFromAttr & pConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, .PeriodValueTo)
                End If
              ElseIf Len(.PeriodValue) > 0 Then
                If .ToAttribute.Length > 0 Then
                  vWhere = vWhere & " AND " & vFromAttr & pConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, .PeriodValue) & " AND " & vToAttr & pConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, .PeriodValue)
                Else
                  vWhere = vWhere & " AND " & vFromAttr & pConn.SQLLiteral("=", CDBField.FieldTypes.cftDate, .PeriodValue)
                End If
              End If

              If Len(vContactWhere) > 0 Or .TableName = "contact_expenditure" Then
                If .FirstSpecial.Length > 0 Then
                  vContactTable = .FirstSpecial & "."
                  vFromSQL = .SpecialTables & .TableName & " WHERE "
                  If vContactWhere.Length > 0 Then vFromSQL = vFromSQL & vContactWhere & " AND "
                  vFromSQL = vFromSQL & .SpecialLink & " AND " & vWhere
                Else
                  vContactTable = .TableName & "."
                  vFromSQL = .TableName
                  If vContactWhere.Length > 0 Then
                    If .TableName = "contact_expenditure" Or .CorO = "C" Then
                      vFromSQL = vFromSQL & " WHERE " & vContactWhere & " AND " & vWhere
                    Else
                      vFromSQL = vFromSQL & " WHERE " & vOrgWhere & " AND " & vWhere
                    End If
                  Else
                    vFromSQL = vFromSQL & " WHERE " & vWhere
                  End If
                End If
              Else
                vContactTable = "ft." 'NoTranslate
                vFromSQL = "contact_header ft, " & .SpecialTables & .TableName & " WHERE "
                If .FirstSpecial.Length > 0 Then
                  vFromSQL = vFromSQL & "ft.contact_number = " & .FirstSpecial & ".contact_number AND " & .SpecialLink & " AND " & vWhere
                Else
                  vFromSQL = vFromSQL & "ft.contact_number = " & .TableName & "." & .ContactAttr & " AND " & vWhere
                End If
                'Here we are always starting with the contact_header table so contact_number is always valid
                .ContactAttr = "contact_number"
              End If
              vAttrs = vContactTable & .ContactAttr & "," & vMaintAttr
              If .SubsidiaryAttribute.Length > 0 Then vAttrs = vAttrs & "," & vSubsidiaryAttr
              If .FromAttribute.Length > 0 Then vAttrs = vAttrs & "," & vFromAttr

              .RecordSet = pConn.GetRecordSet("SELECT " & vAttrs & " FROM " & vFromSQL & " ORDER BY " & vContactTable & .ContactAttr)
              If .RecordSet.Fetch() = False Then .EndOfRecordSet = True
            End If
          End With
        Next

        LogStatus(pJob, (ProjectText.String30341)) 'Scoring Contacts
        While Not vDone
          vRow = 0
          vPoints = 0
          vLastScore = vScoringRows(0).Score
          Do
            With vScoringRows(vRow)
              .NumericValue = 0
              vFound = False
              If .ContactHeader Then .EndOfRecordSet = vCHEOF
              'Now move the cusor up to the current contact number
              If Not .EndOfRecordSet Then
                Do While vContactNumber > .RecordSet.Fields(.ContactAttr).IntegerValue
                  If .RecordSet.Fetch() = False Then
                    .EndOfRecordSet = True
                    If .ContactHeader Then vCHEOF = True
                    Exit Do
                  End If
                Loop
              End If

              If Not .EndOfRecordSet Then
                If vContactNumber = .RecordSet.Fields(.ContactAttr).IntegerValue Then
                  If .ContactHeader Then
                    vFound = False
                    vMain = .RecordSet.Fields(.MainAttribute).Value
                    If .SubsidiaryAttribute.Length > 0 Then vSub = .RecordSet.Fields(.SubsidiaryAttribute).Value
                    If .FromAttribute.Length > 0 Then vFrom = .RecordSet.Fields(.FromAttribute).Value
                    If .ToAttribute.Length > 0 Then vTo = .RecordSet.Fields(.ToAttribute).Value

                    If .MainValueTo.Length > 0 Then 'It's a range
                      If .MainDataType = "C" Then
                        vFound = vMain >= .MainValue And vMain <= .MainValueTo
                      Else
                        vFound = Val(vMain) >= Val(.MainValue) And Val(vMain) <= Val(.MainValueTo)
                      End If
                    Else
                      vFound = IsInList(.MainValue, vMain, .MainDataType)
                    End If
                    If vFound Then
                      If .SubsidiaryAttribute.Length > 0 Then
                        .NumericValue = Val(vSub)
                      Else
                        .NumericValue = Val(vMain)
                      End If
                      If Len(.SubsidiaryValueTo) > 0 Then 'It's a range
                        vFound = vSub >= .SubsidiaryValue And vSub <= .SubsidiaryValueTo
                      ElseIf Len(.SubsidiaryValue) > 0 Then
                        vFound = IsInList(.SubsidiaryValue, vSub, .SubsidiaryDataType)
                      End If
                      If vFound Then
                        If Len(.PeriodValueTo) > 0 Then 'It's a range
                          If .ToAttribute.Length > 0 Then
                            vFound = vFrom <= .PeriodValue And vTo >= .PeriodValueTo
                          Else
                            vFound = vFrom <= .PeriodValue And vFrom >= .PeriodValueTo
                          End If
                        ElseIf Len(.PeriodValue) > 0 Then
                          If .ToAttribute.Length > 0 Then
                            vFound = vFrom <= .PeriodValue And vTo >= .PeriodValue
                          Else
                            vFound = vFrom = .PeriodValue
                          End If
                        End If
                      End If
                    End If
                  Else
                    vFound = True
                    If .ValueRequired Then
                      .NumericValue = 0
                      If .SubsidiaryAttribute.Length > 0 Then
                        .NumericValue = Val(.RecordSet.Fields(.SubsidiaryAttribute).Value)
                      Else
                        If .MainDataType = "I" Or .MainDataType = "N" Then
                          .NumericValue = Val(.RecordSet.Fields(.MainAttribute).Value)
                        Else
                          If .FromAttribute.Length > 0 Then
                            vDate = .RecordSet.Fields(.FromAttribute).Value
                            If IsDate(vDate) Then .NumericValue = DateDiff(Microsoft.VisualBasic.DateInterval.Month, CDate(vDate), Today) + 1
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If

              If .IorE = "E" Then
                If Not vFound Then vPoints = vPoints + .Points
              Else
                If vFound Then
                  If .ValueRequired Then
                    If .Multiplier > 0 Then
                      vPoints = vPoints + (.NumericValue * .Multiplier)
                    Else
                      If .NumericValue * .Multiplier <> 0 Then
                        vPoints = vPoints + (1 / ((.NumericValue * .Multiplier) * -1))
                      End If
                    End If
                  Else
                    vPoints = vPoints + .Points
                  End If
                End If

              End If
              vWrite = False
              vRow = vRow + 1
              If vRow <= UBound(vScoringRows) Then
                If vScoringRows(vRow).Score <> vLastScore Then vWrite = True
                vLastScore = vScoringRows(vRow).Score
              Else
                vWrite = True
              End If
              If vWrite Then
                If vPoints > 9999.99 Then vPoints = 9999.99
                vInsertFields(1).Value = CStr(vContactNumber)
                vInsertFields(2).Value = .Score
                vInsertFields(3).Value = CStr(vPoints)
                pConn.InsertRecord("contact_scores", vInsertFields)
                vPoints = 0
              End If
            End With
          Loop While vRow <= UBound(vScoringRows)

          If vRS.Fetch() = True Then
            vContactNumber = vRS.Fields("contact_number").IntegerValue
          Else
            vDone = True
          End If
          vRecords = vRecords + 1
          pJob.RecordsProcessed = vRecords
        End While

        'Now close the cursors
        For vRow = 0 To UBound(vScoringRows)
          With vScoringRows(vRow)
            If .ContactHeader = True Then
              If vAllContacts And Not vCHRS Is Nothing Then
                vCHRS.CloseRecordSet()
              End If
            Else
              If Not .RecordSet Is Nothing Then
                .RecordSet.CloseRecordSet()
              End If
            End If
          End With
        Next
        vRS.CloseRecordSet()

        If pContactNumber = 0 Then
          LogStatus(pJob, (ProjectText.String30334)) 'Creating Indexes for Contact Scores
          vCDBIndexes.ReCreate(pConn)
        End If
        Return vRecords
      Catch vEx As Exception
        PreserveStackTrace(vEx)
        vCDBIndexes.ReCreate(pConn)
        Throw vEx
      End Try
    End Function
    Private Function BetweenAttrValue(ByRef pDataType As String, ByRef pValue As String, ByRef pValueTo As String) As String
      Select Case pDataType
        Case "I", "L", "N"
          Return " BETWEEN " & pValue & " AND " & pValueTo
        Case Else       ' "C"
          Return " BETWEEN '" & pValue & "' AND '" & pValueTo & "'"
      End Select
    End Function

    Private Function IsInList(ByVal pList As String, ByVal pValue As String, ByVal pDataType As String) As Boolean
      Dim vValues() As String
      Dim vIndex As Integer

      vValues = Split(pList, ",")
      For vIndex = 0 To UBound(vValues)
        If pDataType = "C" Then
          If pValue = vValues(vIndex) Then
            IsInList = True
            Exit For
          End If
        Else
          If Val(pValue) = Val(vValues(vIndex)) Then
            IsInList = True
            Exit For
          End If
        End If
      Next
    End Function

    Private Function InAttrValue(ByVal pDataType As String, ByVal pValue As String) As String
      Dim vValues() As String
      Dim vIndex As Integer
      Dim vResult As String

      Select Case pDataType
        Case "C"
          If InStr(pValue, ",") > 0 Then
            vValues = Split(pValue, ",")
            vResult = " IN ("
            For vIndex = 0 To UBound(vValues)
              If vIndex > 0 Then vResult = vResult & ","
              vResult = vResult & "'" & vValues(vIndex) & "'"
            Next
            Return vResult & ")"
          Else
            Return " = '" & pValue & "'"
          End If
        Case "I", "L", "N"
          If InStr(pValue, ",") > 0 Then
            Return " IN (" & pValue & ")"
          Else
            Return " = " & pValue
          End If
        Case Else                       'Undefined
          Return " = " & pValue
      End Select
    End Function

    Private Function ParseValue(ByRef pValue As String, ByRef pFromValue As String, ByRef pToValue As String) As Boolean
      Dim vPos As Integer
      Dim vInquotes As String = ""
      Dim vChar As String
      Dim vWord As String = ""
      Dim vEndofWord As Boolean
      Dim vToFound As Boolean
      Dim vLen As Integer
      Dim vError As Boolean

      vLen = pValue.Length
      For vPos = 1 To vLen
        vChar = Mid(pValue, vPos, 1)
        Select Case vChar
          Case "'", Chr(34)
            If vWord.Length = 0 Then
              vInquotes = vChar
            ElseIf vInquotes = vChar Then
              vInquotes = ""
              vEndofWord = True
            Else
              vWord = vWord & vChar
            End If
          Case " ", ","
            If vInquotes = "" Then
              vEndofWord = True
            Else
              vWord = vWord & vChar
            End If
          Case Else
            vWord = vWord & vChar
        End Select
        If vEndofWord Or vPos = vLen Then
          vWord = Trim(vWord)
          If vWord.Length > 0 Then
            If vToFound Then
              If pToValue.Length > 0 Then
                vError = True
              Else
                pToValue = vWord
              End If
            Else
              If UCase(vWord) = "TO" Then
                If InStr(pFromValue, ",") > 0 Then
                  vError = True
                Else
                  vToFound = True
                End If
              Else
                If pFromValue.Length > 0 Then
                  pFromValue = pFromValue & "," & vWord
                Else
                  pFromValue = vWord
                End If
              End If
            End If
            vWord = ""
          End If
          vEndofWord = False
        End If
      Next
      ParseValue = vError
    End Function

    Public Property LogFile() As LogFile
      Get
        Return mvLogFile
      End Get
      Set(ByVal Value As LogFile)
        mvLogFile = Value
      End Set
    End Property

    Private Sub LogStatus(ByVal pJob As JobSchedule, ByVal pMsg As String)
      pJob.InfoMessage = pMsg
      If Not mvLogFile Is Nothing Then mvLogFile.WriteLine(pMsg & " " & TodaysDateAndTime())
    End Sub
  End Class
End Namespace

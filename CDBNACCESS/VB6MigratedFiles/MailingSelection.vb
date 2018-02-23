Imports System.IO

Namespace Access
  Public Class MailingSelection

    Private Const MAX_RECORDS_IN_GRID As Integer = 65535          'These constants are also set on the client in the ListManager class
    Private Const FILTERED_RECORD_LIMIT As Integer = 1000         'These constants are also set on the client in the ListManager class
    Private Const UNFILTERED_RECORD_LIMIT As Integer = 100        'These constants are also set on the client in the ListManager class

    Public Enum SelectionMergeMode
      smmNone
      smmAAndB 'A and B
      smmAOrB 'A or B
      smmBAndNotInA 'B and not in A
      smmNotInAAndB 'not in A and B
      smmAAndNotInB 'A and not in B
    End Enum

    Public Enum MailingTypes
      mtyGeneralMailing = 1
      mtyDirectDebits
      mtyStandingOrders
      mtyMembers
      mtyMembershipCards
      mtySubscriptions
      mtyPayers
      mtySelectionTester
      mtyCampaigns
      mtyScoringAnalyser
      mtyPerformanceAnalyser
      mtyNameGathering
      mtyRenewalsAndReminders
      mtyStandardExclusions
      mtyStandingOrderCancellation
      mtyEventBookings
      mtyEventAttendees
      mtyEventPersonnel
      mtyMemberFulfilment
      mtyNonMemberFulfilment
      mtySaleOrReturn
      mtyGAYECancellation
      mtyGAYEPledges
      mtyEventSponsors
      mtyIrishGiftAid
      mtyExamBookings
      mtyExamCandidates
      mtyExamCertificates
      'Put any other enum values ABOVE this comment
      mtyMaxMailingType
    End Enum

    Public Enum MailingGenerateResult
      mgrNone = 1
      mgrRefine
      mgrReset
    End Enum

    Public Enum SortOrderTypes
      sotBranch
      sotCountry
      sotMailsort
      sotSurname
      sotOther1
      sotOther2
    End Enum

    Public NewCriteriaSet As Integer 'Selected from the Criteria Sets List
    Public NewSelectionSet As Integer 'Selected from the Selection Sets List

    Public GenerateStatus As MailingGenerateResult
    Public Count As Integer
    Public CriteriaRows As Integer 'Count of current rows
    Public CriteriaSetNumber As Integer 'Current Criteria Set
    Public SelectionSetNumber As Integer 'Current Selection Set
    Public Revision As Integer
    Public ExclusionCriteriaSet As Integer

    Public BypassCriteriaCount As Boolean 'Used by Selection Manager to indicate whether each criteria line should be counted before the main SQL is built.
    'This replicates the existing bypass-count facility that's available within appeal mailings.

    Private Enum TokenActionTypes
      tatParseOnly = 1
      tatSubstitute
      tatLocate
    End Enum

    Private Enum CriteriaItems
      citMainValue = 1
      citSubValue
      citPeriod
    End Enum

    Public Enum MembershipCardProductionTypes
      mcpDefault = 1
      mcpAutoOrPaid
      mcpPaymentRequired
    End Enum

    Public Enum OrgSelectContact
      oscNone = 0
      oscAllEmployees = 1
      oscDefaultContact
      oscOrganisation
    End Enum

    Public Enum OrgSelectAddress
      osaNone = 0
      osaOrganisationAddress = 1
      osaDefaultAddress
      osaAddressByUsage
    End Enum

    Private Enum BulkEmailSelectionTypes
      bestByUsage = 1
      bestByPreferred
      bestByDefault
      bestByAny
    End Enum

    Private mvEnv As CDBEnvironment
    Private mvConn As CDBConnection
    Private mvApplication As String
    Private mvPosition As Integer
    Private mvToken1 As String
    Private mvToken2 As String
    Private mvTokenDelimiter As String
    Private mvContactAddressTable As String
    Private mvOrgAddressTable As String
    Private mvMasterTable As String
    Private mvOrgTable As String
    Private mvPositionTable As String
    Private mvAddressAttribute As String
    Private mvContactAttribute As String
    Private mvMasterAttribute As String
    Private mvOrgAttribute As String
    Private mvWhere As String
    Private mvTableList As String
    Private mvSingleWhere As String
    Private mvSingleTableList As String
    Private mvTableNumber As Integer
    Private mvLastCTable As String
    Private mvLastOTable As String
    Private mvLastCSearchArea As String
    Private mvLastOSearchArea As String
    Private mvCGeographic As String
    Private mvOGeographic As String
    Private mvCAddressLink As String
    Private mvOAddressLink As String
    Private mvContactTableAlias As String
    Private mvOrgTableAlias As String
    Private mvIndexNoOptimise As Boolean
    Private mvConCriteriaCount As Integer
    Private mvOrgCriteriaCount As Integer
    Private mvMailingType As MailingTypes
    Private mvSelectionTable As String
    Private mvCaption As String
    Private mvCardProductionType As MembershipCardProductionTypes
    Private mvOrgRoles As String
    Private mvOrgMailTo As OrgSelectContact
    Private mvOrgMailWhere As OrgSelectAddress
    Private mvOrgAddUsage As String
    Private mvOrgLabelName As String
    Private mvVariableParameters As String = ""
    Private mvCurrentCriteria As CriteriaDetails
    Private mvSelectionSetTable As String
    Private mvTempTableName As String
    Private mvAppealMailing As Boolean
    Private mvLinkToContacts As Boolean
    Private mvSegmentScoreOrRandom As Boolean
    Private mvDisplayOrgSelection As Boolean
    Private mvIncludeHistoricRoles As Boolean
    Private mvMasterHasAddress As Boolean
    Private mvMasterTableAlias As String

    Private mvVariableCriteria As Collection
    Private mvCriteriaContexts As Collection

    Private mvDestinationAttributes As String
    Private mvSelectionAttributes As String

    Private mvTableAliases As CDBCollection
    Private mvMasterAttrTables As CDBCollection
    Private mvOrgAttrTables As CDBCollection

    Private mvLMMaxContact As Integer 'List Manager max contact number for oracle

    Private mvOrgMailOptionsChanged As Boolean
    Private mvSegment As Segment
    Private mvSegmentsHaveSelectionOptions As Boolean

    Private mvExamUnitCertRunType As ExamUnitCertRunType = Nothing
    '
    Private Sub ConvertToJoints(ByVal pTable As String, ByVal pStatuses As String, ByRef pAppeal As Appeal, ByVal pSequence As Integer)
      Dim vRealToJoint As String
      Dim vDerivedToJoint As String
      Dim vSQL As String
      Dim vRowsAffected As Integer
      Dim vBaseSQL As String

      vRealToJoint = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink)
      vDerivedToJoint = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink)

      mvConn.LogMailSQL("Starting Joint Conversion")

      'Build a temporary table containing the joint contact information for those selected contacts that are in a joint relationship
      'Because of performance reasons we first are going to create an _6 temporary table with the joints in it

      vSQL = "INSERT INTO " & pTable & "_6 (" & mvDestinationAttributes & ",joint_contact_number,joint_address_number)"
      vSQL = vSQL & " SELECT y.segment_sequence, y.selection_set, y.revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & "y." & mvMasterAttribute & ", "
      vSQL = vSQL & "y.contact_number, y.address_number, y.address_number_2,"
      vSQL = vSQL & "c.contact_number AS joint_contact_number, c.address_number AS joint_address_number"
      vSQL = vSQL & " FROM " & pTable & " y"
      vSQL = vSQL & " INNER JOIN contact_links cl ON y.contact_number = cl.contact_number_2"
      vSQL = vSQL & " INNER JOIN contacts c ON cl.contact_number_1 = c.contact_number AND c.address_number = y.address_number"
      If pAppeal.MailJointsMethod = Appeal.MailJointMethods.mjmBothContactsSelected Then
        vSQL = vSQL & " INNER JOIN contact_links cl2 ON cl.contact_number_1 = cl2.contact_number_1"
        vSQL = vSQL & " INNER JOIN " & pTable & " z ON cl2.contact_number_2 = z.contact_number AND z.contact_number <> y.contact_number"
      End If
      vSQL = vSQL & " WHERE"
      If vRealToJoint = vDerivedToJoint Then
        vSQL = vSQL & " cl.relationship = '" & vRealToJoint & "'"
      Else
        vSQL = vSQL & " cl.relationship IN ('" & vRealToJoint & "','" & vDerivedToJoint & "')"
      End If
      vSQL = vSQL & " AND (cl.historical IS NULL OR cl.historical = 'N')"
      If Len(pStatuses) > 0 Then vSQL = vSQL & " AND (c.status IS NULL OR (c.status IS NOT NULL AND c.status NOT IN (" & pStatuses & ")))"
      If pAppeal.MailJointsMethod = Appeal.MailJointMethods.mjmBothContactsSelected Then
        If vRealToJoint = vDerivedToJoint Then
          vSQL = vSQL & " AND cl2.relationship = '" & vRealToJoint & "'"
        Else
          vSQL = vSQL & " AND cl2.relationship IN ('" & vRealToJoint & "','" & vDerivedToJoint & "')"
        End If
        vSQL = vSQL & " AND (cl2.historical IS NULL OR cl2.historical = 'N')"
      End If
      mvConn.LogMailSQL(vSQL)
      vRowsAffected = mvConn.ExecuteSQL(mvConn.ProcessAnsiJoins(vSQL))

      vSQL = "INSERT INTO " & pTable & "_3 (" & mvDestinationAttributes & ",joint_contact_number,joint_address_number)"
      vSQL = vSQL & " SELECT x.segment_sequence, x.selection_set, x.revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & "x." & mvMasterAttribute & ", "
      vSQL = vSQL & "x.contact_number, x.address_number, x.address_number_2, jc.joint_contact_number, jc.address_number AS joint_address_number"
      vSQL = vSQL & " FROM " & pTable & " x"
      vSQL = vSQL & " LEFT OUTER JOIN " & pTable & "_6 jc ON x.contact_number = jc.contact_number"
      mvConn.LogMailSQL(vSQL)
      vRowsAffected = mvConn.ExecuteSQL(mvConn.ProcessAnsiJoins(vSQL))

      'Only need to continue if at least one record can be converted
      If vRowsAffected > 0 Then
        'If conversion method is not Always then need to further refine those records that are marked for conversion
        Select Case pAppeal.MailJointsMethod
          Case Appeal.MailJointMethods.mjmBothContactsSelected, Appeal.MailJointMethods.mjmOneSelectedOneNotExcluded
            'Create index on the temporary table containing the joint contact information
            mvConn.CreateIndex(False, pTable & "_3", {"joint_contact_number", "selection_set"})
            'Remove any joint contact information from those records where the joint contact number appears in the exclusion segment
            vSQL = "UPDATE " & pTable & "_3 SET joint_contact_number = NULL, joint_address_number = NULL"
            vSQL = vSQL & " WHERE joint_contact_number IN (SELECT x.joint_contact_number FROM " & pTable & "_3 x WHERE x.selection_set = 0 AND x.joint_contact_number IS NOT NULL)"
            mvConn.LogMailSQL(vSQL)
            vRowsAffected = mvConn.ExecuteSQL(vSQL)

            If pAppeal.MailJointsMethod = Appeal.MailJointMethods.mjmBothContactsSelected Then
              'Create index on the temporary table containing the joint contact information
              mvConn.CreateIndex(False, pTable & "_3", {"contact_number"})
              'Remove any row where the joint contact appears in an earlier segment, but not excluded
              vSQL = "DELETE FROM " & pTable & "_3 WHERE contact_number IN"
              vSQL = vSQL & " (SELECT x.contact_number FROM " & pTable & "_3 x"
              vSQL = vSQL & " INNER JOIN contact_links cl ON x.contact_number = cl.contact_number_2"
              vSQL = vSQL & " INNER JOIN contacts c ON cl.contact_number_1 = c.contact_number AND c.address_number = x.address_number"
              vSQL = vSQL & " INNER JOIN " & pTable & "_3 y ON c.contact_number = y.contact_number"
              vSQL = vSQL & " WHERE cl.relationship "
              If vRealToJoint = vDerivedToJoint Then
                vSQL = vSQL & " = '" & vRealToJoint & "'"
              Else
                vSQL = vSQL & " IN ('" & vRealToJoint & "','" & vDerivedToJoint & "')"
              End If
              vSQL = vSQL & " AND (cl.historical IS NULL OR cl.historical = 'N')"
              If Len(pStatuses) > 0 Then vSQL = vSQL & " AND (c.status IS NULL OR (c.status IS NOT NULL AND c.status NOT IN (" & pStatuses & ")))"
              vSQL = vSQL & " AND y.segment_sequence < x.segment_sequence AND y.selection_set > 0)"
              If pSequence > 0 Then vSQL = vSQL & " AND segment_sequence = " & pSequence
              mvConn.LogMailSQL(vSQL)
              vRowsAffected = mvConn.ExecuteSQL(mvConn.ProcessAnsiJoins(vSQL))
            End If
        End Select
        'Perform the actual joint conversion by building another temporary table containing the correct contact and address information of the selected contacts
        vBaseSQL = "INSERT INTO " & pTable & "_4 (" & mvDestinationAttributes & ")"
        vBaseSQL = vBaseSQL & " SELECT x.segment_sequence, x.selection_set, x.revision, "
        If mvMasterAttribute <> mvContactAttribute Then vBaseSQL = vBaseSQL & "x." & mvMasterAttribute & ", "
        vBaseSQL = vBaseSQL & "%1, x.address_number_2"
        vBaseSQL = vBaseSQL & " FROM " & pTable & "_3 x WHERE joint_contact_number %2"
        'Firstly, the contact & address numbers of those selected contacts that have not been converted
        vSQL = Replace(vBaseSQL, "%1", "x.contact_number, x.address_number")
        vSQL = Replace(vSQL, "%2", "IS NULL")
        mvConn.LogMailSQL(vSQL)
        vRowsAffected = mvConn.ExecuteSQL(vSQL)
        'Next, the joint contact & address numbers of those selected contacts that have been converted
        vSQL = Replace(vBaseSQL, "%1", "x.joint_contact_number, x.joint_address_number")
        vSQL = Replace(vSQL, "%2", "IS NOT NULL")
        mvConn.LogMailSQL(vSQL)
        vRowsAffected = mvConn.ExecuteSQL(vSQL)
        'Reset this variable so that the deduplication routine processes the correct table
        mvTempTableName = pTable & "_4"
        'Now need to add the same indices to this table as the original mvTempTableName has
        If mvMasterAttribute <> mvContactAttribute Then
          mvConn.CreateIndex(False, mvTempTableName, {"selection_set", mvMasterAttribute, "contact_number", "address_number"})
        Else
          mvConn.CreateIndex(False, mvTempTableName, {"selection_set", "contact_number", "address_number"})
        End If
        mvConn.CreateIndex(False, mvTempTableName, {"segment_sequence", "contact_number"})
        mvConn.CreateIndex(False, mvTempTableName, {"contact_number"})
      End If
      mvConn.LogMailSQL("Joint Conversion Completed")
    End Sub
    Private Sub ProcessRoles(ByVal pTable As String, Optional ByVal pSelectionSet As Integer = 0)
      Dim vRowsAffected As Integer
      Dim vCurrentContact As Integer
      Dim vCurrentAddress As Integer
      Dim vCurrentRole As String = ""
      Dim vSelectedContact As Integer
      Dim vSelectedAddress As Integer
      Dim vNewContact As Integer
      Dim vNewAddress As Integer
      Dim vRole As String
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields

      With vUpdateFields
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("address_number_2", CDBField.FieldTypes.cftLong)
      End With
      With vWhereFields
        If pSelectionSet > 0 Then .Add("selection_set", CDBField.FieldTypes.cftLong, pSelectionSet)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
      End With

      vSQL = "SELECT x.contact_number AS x_contact, x.address_number AS x_address, cp.contact_number, cp.address_number, cr.role"
      vSQL = vSQL & " FROM " & pTable & " x ,contact_roles cr, contact_positions cp WHERE"
      If pSelectionSet > 0 Then vSQL = vSQL & " x.selection_set = " & pSelectionSet & " AND"
      vSQL = vSQL & " x.contact_number = cr.organisation_number AND role IN (" & mvOrgRoles & ")"
      If Not mvIncludeHistoricRoles Then vSQL = vSQL & " AND cr.is_active = 'Y'"
      vSQL = vSQL & " AND cr.organisation_number = cp.organisation_number AND cr.contact_number = cp.contact_number AND " & mvConn.DBSpecialCol("cp", "current") & " = 'Y'"
      vRecordSet = mvConn.GetRecordSet(vSQL)

      While vRecordSet.Fetch() = True
        vSelectedContact = vRecordSet.Fields("x_contact").IntegerValue
        vSelectedAddress = vRecordSet.Fields("x_address").IntegerValue
        vRole = vRecordSet.Fields("role").Value
        If vSelectedContact = vCurrentContact And vSelectedAddress = vCurrentAddress Then
          If InStr(mvOrgRoles, vRole) < InStr(mvOrgRoles, vCurrentRole) Then
            'Better role then previous role so hold info
            vCurrentRole = vRole
            vNewContact = vRecordSet.Fields("contact_number").IntegerValue
            vNewAddress = vRecordSet.Fields("address_number").IntegerValue
          End If
        Else
          If vCurrentContact > 0 Then 'If info held from previous time then update selected contacts
            With vUpdateFields
              .Item("contact_number").Value = CStr(vNewContact)
              .Item("address_number").Value = CStr(vNewAddress)
              .Item("address_number_2").Value = CStr(vNewAddress)
            End With
            With vWhereFields
              .Item("contact_number").Value = CStr(vCurrentContact)
              .Item("address_number").Value = CStr(vCurrentAddress)
            End With
            vRowsAffected = mvConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
            If vRowsAffected = 0 Then vRowsAffected = mvConn.DeleteRecords(pTable, vWhereFields, False) 'Assume duplicate !!!!  TODO
          End If
          vCurrentContact = vSelectedContact
          vCurrentAddress = vSelectedAddress
          vCurrentRole = vRole
          vNewContact = vRecordSet.Fields("contact_number").IntegerValue
          vNewAddress = vRecordSet.Fields("address_number").IntegerValue
        End If
      End While
      vRecordSet.CloseRecordSet()
      'if one left do the update
      If vCurrentContact > 0 Then 'If info held from previous time then update selected contacts
        With vUpdateFields
          .Item("contact_number").Value = CStr(vNewContact)
          .Item("address_number").Value = CStr(vNewAddress)
          .Item("address_number_2").Value = CStr(vNewAddress)
        End With
        With vWhereFields
          .Item("contact_number").Value = CStr(vCurrentContact)
          .Item("address_number").Value = CStr(vCurrentAddress)
        End With
        vRowsAffected = mvConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
        If vRowsAffected = 0 Then vRowsAffected = mvConn.DeleteRecords(pTable, vWhereFields, False) 'Assume duplicate !!!!  TODO
      End If
    End Sub
    Public Function ProcessSelection(ByRef pCriteriaSet As Integer, ByRef pSelectionSet As Integer, ByRef pRevision As Integer, ByRef pMode As SelectionMergeMode) As Integer
      Dim vCount As Integer
      Dim vRevision As Integer
      Dim vStatement As String
      Dim vAttributes As String
      Dim vInsert As String
      Dim vInsertTemp As String
      Dim vAddrTable As String
      Dim vSelectTable As String
      Dim vTempTable As String
      Dim vDedupTable As String
      Dim vProcessRoles As Boolean
      Dim vDedupTableOnly As Boolean

      vTempTable = mvSelectionTable & Mid("_temp", 1, 28 - Len(mvSelectionTable))
      vDedupTable = mvSelectionTable & Mid("_dedup", 1, 30 - Len(mvSelectionTable))
      vProcessRoles = (mvOrgMailWhere > 0 And mvOrgMailTo = OrgSelectContact.oscOrganisation And mvOrgRoles <> "")
      vDedupTableOnly = (pMode = SelectionMergeMode.smmAAndNotInB And Not vProcessRoles)

      vCount = RoughCount(pCriteriaSet)
      If vCount = 0 Then RaiseError(DataAccessErrors.daeNoContactsFound)
      ClearTableAliases(True)
      BuildStatement(pCriteriaSet)
      vRevision = pRevision + 1
      vAttributes = "selection_set,revision,"
      If mvMasterAttribute <> mvContactAttribute Then
        vAttributes = vAttributes & mvMasterAttribute & ","
      End If
      vAttributes = vAttributes & mvContactAttribute & ", address_number, address_number_2"
      vInsert = "INSERT INTO " & mvSelectionTable & " (" & vAttributes & ")"
      vInsertTemp = "INSERT INTO " & If(vDedupTableOnly, vDedupTable, vTempTable) & " (" & vAttributes & ")"

      'create two work tables
      If Not vDedupTableOnly Then CreateTempTables(False, mvSelectionTable, vTempTable, True, True)
      CreateTempTables(False, mvSelectionTable, vDedupTable, True)

      Select Case pMode
        Case SelectionMergeMode.smmAAndB
          If Not vProcessRoles Then
            mvTableNumber = mvTableNumber + 1
            vSelectTable = "table" & CStr(mvTableNumber)
            mvTableList = mvTableList & ", " & mvSelectionTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvSelectionTable, vSelectTable)
            mvWhere = mvWhere & " AND " & mvMasterTableAlias & "." & mvMasterAttribute & " = " & vSelectTable & "." & mvMasterAttribute & " AND " & vSelectTable & "." & "selection_set = abs(" & pSelectionSet.ToString & ") AND " & vSelectTable & "." & "revision = " & pRevision.ToString
          End If
        Case SelectionMergeMode.smmBAndNotInA, SelectionMergeMode.smmAOrB
          If pMode = SelectionMergeMode.smmAOrB Or (pMode = SelectionMergeMode.smmBAndNotInA And Not vProcessRoles) Then
            mvTableNumber = mvTableNumber + 1
            vSelectTable = "table" & CStr(mvTableNumber)
            mvWhere = mvWhere & " AND " & mvMasterTableAlias & "." & mvMasterAttribute & " NOT IN (SELECT " & mvMasterAttribute & " FROM " & mvSelectionTable & " " & vSelectTable & " WHERE " & vSelectTable & "." & mvMasterAttribute & " = " & mvLastCTable & "." & mvMasterAttribute & " AND " & vSelectTable & "." & "selection_set = abs(" & pSelectionSet.ToString & ") AND " & vSelectTable & "." & "revision = " & pRevision.ToString & ")"
          End If
        Case SelectionMergeMode.smmNotInAAndB
          ' nothing to do
        Case SelectionMergeMode.smmAAndNotInB
          ' nothing to do
        Case Else
          ' nothing to do
      End Select

      If pMode = SelectionMergeMode.smmNone Or Not vDedupTableOnly Then
        vAddrTable = LinkToAddress()
        vStatement = vInsertTemp & " SELECT " & pSelectionSet.ToString & ", " & vRevision.ToString
        If mvMasterAttribute <> mvContactAttribute Then
          vStatement = vStatement & ", "
          If Not vProcessRoles Then vStatement = vStatement & mvMasterTableAlias & "."
          vStatement = vStatement & mvMasterAttribute
        End If
        'Here we need to add the table name for the contact attribute since there may be more than one
        vStatement = vStatement & ", " & mvLastCTable & "." & mvContactAttribute & ", " & vAddrTable & "." & mvAddressAttribute
        If InStr(mvTableList, mvSelectionTable) > 0 Then
          vStatement = vStatement & ", address_number_2"
        Else
          vStatement = vStatement & ", " & vAddrTable & "." & mvAddressAttribute
        End If
        vStatement = vStatement & " FROM " & mvTableList & mvWhere
      Else
        mvTableNumber = mvTableNumber + 1
        vSelectTable = "table" & CStr(mvTableNumber)
        vStatement = vInsertTemp & " SELECT " & pSelectionSet.ToString & ", " & vRevision.ToString
        If mvMasterAttribute <> mvContactAttribute Then
          vStatement = vStatement & ", " & mvMasterAttribute
        End If
        vStatement = vStatement & ", " & mvContactAttribute & ", " & mvAddressAttribute & ", address_number_2" & " FROM " & mvSelectionTable & " " & vSelectTable & " WHERE " & vSelectTable & "." & "selection_set = " & pSelectionSet.ToString & " AND " & vSelectTable & "." & "revision = " & pRevision.ToString & " AND " & vSelectTable & "." & mvMasterAttribute & " NOT IN (SELECT " & mvMasterTableAlias & "." & mvMasterAttribute & " FROM " & mvTableList & mvWhere & ") "
      End If

      'Write Refinement SQL Log Header
      mvConn.LogMailSQL("***** User : " & mvEnv.User.Logname & ", Date : " & Now.ToString & ", Application : " & mvCaption & " - Refinement *****")

      'add any records that match the entered criteria to the work table
      mvConn.LogMailSQL(vStatement)
      mvConn.ExecuteSQL(vStatement)
      If Not vDedupTableOnly Then
        'create index on work table
        If mvMasterAttribute <> mvContactAttribute Then
          mvConn.CreateIndex(False, vTempTable, {"selection_set", mvMasterAttribute, "contact_number", "address_number"})
        Else
          mvConn.CreateIndex(False, vTempTable, {"selection_set", "contact_number", "address_number"})
        End If
        'check if we should replace some of the contact information
        If vProcessRoles Then ProcessRoles(vTempTable)
        'dedup the records in the work table into the dedup table
        DedupRecords(vDedupTable, vTempTable)
      End If
      mvConn.CreateIndex(True, vDedupTable, {"selection_set", "revision", mvMasterAttribute})
      'drop the work table
      If Not vDedupTableOnly Then mvConn.ExecuteSQL("DROP TABLE " & vTempTable)

      Select Case pMode
        Case SelectionMergeMode.smmBAndNotInA, SelectionMergeMode.smmAAndNotInB
          If vProcessRoles Then
            CreateTempTables(False, mvSelectionTable, vTempTable, True)
            vStatement = vInsertTemp & " SELECT " & pSelectionSet.ToString & ", " & vRevision.ToString
            If mvMasterAttribute <> mvContactAttribute Then vStatement = vStatement & ", " & mvMasterAttribute
            vStatement = vStatement & ", " & mvContactAttribute & ", " & mvAddressAttribute & ", address_number_2"
            vStatement = vStatement & " FROM {0}"
            vStatement = vStatement & " WHERE " & mvMasterAttribute & " NOT IN (SELECT " & mvMasterAttribute & " FROM {1})"
            If pMode = SelectionMergeMode.smmBAndNotInA Then
              mvConn.LogMailSQL(XLATP2(vStatement, vDedupTable, mvSelectionTable))
              mvConn.ExecuteSQL(String.Format(vStatement, vDedupTable, mvSelectionTable))
            Else
              mvConn.LogMailSQL(XLATP2(vStatement, mvSelectionTable, vDedupTable))
              mvConn.ExecuteSQL(XLATP2(vStatement, mvSelectionTable, vDedupTable))
            End If
            CopyDetails(pSelectionSet, 0, pSelectionSet, vRevision, False, False, vTempTable)
            mvConn.ExecuteSQL("DROP TABLE " & vTempTable)
          Else
            CopyDetails(pSelectionSet, 0, pSelectionSet, vRevision, False, False, vDedupTable)
          End If
        Case SelectionMergeMode.smmAAndB
          If vProcessRoles Then
            CreateTempTables(False, mvSelectionTable, vTempTable, True)
            vStatement = vInsertTemp & " SELECT " & pSelectionSet.ToString & ", " & vRevision.ToString
            If mvMasterAttribute <> mvContactAttribute Then vStatement = vStatement & ", y." & mvMasterAttribute
            vStatement = vStatement & ", y." & mvContactAttribute & ", y." & mvAddressAttribute & ", y.address_number_2"
            vStatement = vStatement & " FROM " & mvSelectionTable & " x, " & vDedupTable & " y"
            vStatement = vStatement & " WHERE x." & mvMasterAttribute & " = y." & mvMasterAttribute
            mvConn.LogMailSQL(vStatement)
            mvConn.ExecuteSQL(vStatement)
            CopyDetails(pSelectionSet, 0, pSelectionSet, vRevision, False, False, vTempTable)
            mvConn.ExecuteSQL("DROP TABLE " & vTempTable)
          Else
            CopyDetails(pSelectionSet, 0, pSelectionSet, vRevision, False, False, vDedupTable)
          End If
        Case SelectionMergeMode.smmAOrB
          CopyDetails(pSelectionSet, pRevision, pSelectionSet, vRevision, False, False, mvSelectionTable, vDedupTable)
          CopyDetails(pSelectionSet, 0, pSelectionSet, vRevision, False, False, vDedupTable)
        Case SelectionMergeMode.smmNotInAAndB
          CreateTempTables(False, mvSelectionTable, vTempTable, True)
          vStatement = vInsertTemp & " SELECT " & pSelectionSet.ToString & ", " & vRevision.ToString
          If mvMasterAttribute <> mvContactAttribute Then vStatement = vStatement & ", " & mvMasterAttribute
          vStatement = vStatement & ", " & mvContactAttribute & ", " & mvAddressAttribute & ", address_number_2"
          vStatement = vStatement & " FROM {0} WHERE " & mvMasterAttribute
          vStatement = vStatement & " NOT IN (SELECT " & mvMasterAttribute & " FROM {1} )"
          mvConn.LogMailSQL(XLATP2(vStatement & " AND revision = " & pRevision, mvSelectionTable, vDedupTable))
          mvConn.ExecuteSQL(String.Format(vStatement & " AND revision = " & pRevision, mvSelectionTable, vDedupTable))
          mvConn.LogMailSQL(XLATP2(vStatement, vDedupTable, mvSelectionTable))
          mvConn.ExecuteSQL(String.Format(vStatement, vDedupTable, mvSelectionTable))
          CopyDetails(pSelectionSet, 0, pSelectionSet, vRevision, False, False, vTempTable)
          mvConn.ExecuteSQL("DROP TABLE " & vTempTable)
        Case Else
          ' nothing to do
      End Select

      'drop the dedup table
      mvConn.ExecuteSQL("DROP TABLE " & vDedupTable)

      'Write SQL Log Footer
      mvConn.LogMailSQL("***** Date : " & Now.ToString)

      ProcessSelection = vRevision
    End Function
    Private Sub ScoreContacts(ByRef pSegment As Segment, ByVal pTable As String)
      Dim vRecordSet As CDBRecordSet
      Dim vCount As Integer
      Dim vExclude As Boolean
      Dim vUpdateFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vUpdate As Boolean

      mvConn.LogMailSQL("Starting the Selection of Contacts based on Score")

      If pSegment.RequiredCount > (pSegment.ActualCount - pSegment.RequiredCount) Then
        'you're including more than you're excluding so only mark those that should be excluded
        vExclude = True
      Else
        'you're including less than you're excluding so only mark those that should be included
      End If

      vUpdateFields = New CDBFields
      vUpdateFields.Add("marker", CDBField.FieldTypes.cftCharacter, "X")

      vRecordSet = mvConn.GetRecordSet("SELECT selection_set, x." & mvMasterAttribute & " FROM " & pTable & " x, contact_scores cs WHERE selection_set = " & pSegment.SelectionSet & " AND revision = 1 AND x.contact_number = cs.contact_number AND score = '" & pSegment.Score & "' ORDER BY cs.points DESC")
      With vRecordSet
        While .Fetch() = True
          vWhereFields = New CDBFields
          vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, .Fields(1).IntegerValue)
          vWhereFields.Add("revision", CDBField.FieldTypes.cftInteger, "1")
          vWhereFields.Add(mvMasterAttribute, CDBField.FieldTypes.cftLong, .Fields(2).IntegerValue)
          vUpdate = False
          vCount = vCount + 1
          If vCount > pSegment.RequiredCount Then
            'we've exceeded the req'd count, so if we're excluding mark this record for exclusion
            If vExclude Then vUpdate = True
          Else
            'we're still w/in the req'd count, so if we're including mark this record for inclusion
            If Not vExclude Then vUpdate = True
          End If
          If vUpdate Then mvConn.UpdateRecords(pTable, vUpdateFields, vWhereFields)
        End While
        .CloseRecordSet()
      End With

      vWhereFields = New CDBFields
      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSegment.SelectionSet)
      If vExclude Then
        vWhereFields.Add("marker", CDBField.FieldTypes.cftCharacter, "X")
      Else
        vWhereFields.Add("marker", CDBField.FieldTypes.cftCharacter)
      End If
      mvConn.DeleteRecords(pTable, vWhereFields)

      pSegment.ActualCount = pSegment.RequiredCount

      mvConn.LogMailSQL("Score-based Contact Selection Completed")
    End Sub

    Private Sub SelectContacts(ByRef pCriteriaSet As CriteriaSet, ByRef pCriteriaSets As Collection, ByRef pAppeal As Appeal, ByRef pSegment As Segment, ByVal pTableName As String)
      Dim vCount As Integer
      Dim vSQL As String
      Dim vAttrList As String = ""
      Dim vTableList As String
      Dim vWhere As String = ""
      Dim vCriteriaSet As String
      Dim vAddrTable As String
      Dim vWhereFields As CDBFields
      Dim vDestAttrs As String = ""
      Dim vMailingAccess As String
      Dim vPos As Integer
      Dim vCTable As String
      Dim vStep As SelectionStep

      If mvAppealMailing Then vDestAttrs = "segment_sequence,"
      vDestAttrs = vDestAttrs & "selection_set,revision"
      If mvMasterAttribute <> mvContactAttribute Then vDestAttrs = vDestAttrs & "," & mvMasterAttribute
      vDestAttrs = vDestAttrs & "," & mvContactAttribute & ", address_number, address_number_2"
      If mvMailingType = MailingTypes.mtyIrishGiftAid Then vDestAttrs = vDestAttrs & ",performance"

      If pCriteriaSet.SelectionSteps.Count() > 0 Then
        If mvAppealMailing Then vAttrList = pSegment.SegmentSequence & ","
        vAttrList = vAttrList & pSegment.SelectionSet & ",1,"
        If mvMasterAttribute <> mvContactAttribute Then vAttrList = vAttrList & "x." & mvMasterAttribute & ","
        vAttrList = vAttrList & "x." & mvContactAttribute & "," & "x." & mvAddressAttribute & "," & "x." & mvAddressAttribute 'use the mvAddressAttribute twice to populate both address_number & address_number_2 attributes
        vTableList = mvMasterTable & " x"
        For Each vStep In pCriteriaSet.SelectionSteps
          ProcessStep(pTableName, vDestAttrs, vAttrList, pSegment, vStep, mvAppealMailing, True)
        Next vStep
      Else
        If pCriteriaSets.Count() = 0 Then 'Select ALL records
          If mvAppealMailing Then vAttrList = pSegment.SegmentSequence & ", "
          vAttrList = vAttrList & pSegment.SelectionSet & ", 1, "
          If mvMasterAttribute <> mvContactAttribute Then vAttrList = vAttrList & "x." & mvMasterAttribute & ", "
          vAttrList = vAttrList & "x." & mvContactAttribute & ", " & "x." & mvAddressAttribute & ", " & "x." & mvAddressAttribute 'use the mvAddressAttribute twice to populate both address_number & address_number_2 attributes
          '      If mvMailingType = mtyIrishGiftAid Then vAttrList = vAttrList & ", " & mvMasterTable & ".performance"
          vTableList = mvMasterTable & " x"
          If pSegment.Score.Length > 0 Then
            vTableList = vTableList & ", contact_scores cs"
            vWhere = " WHERE x." & mvContactAttribute & " = cs." & mvContactAttribute & " AND cs.score = '" & pSegment.Score & "'"
          End If
          vSQL = "INSERT INTO " & pTableName & "(" & vDestAttrs & ") SELECT " & vAttrList & " FROM " & vTableList & vWhere
          mvConn.LogMailSQL(vSQL)
          vCount = mvConn.ExecuteSQL(vSQL)
        Else 'Process criteria set(s) for segment
          For Each vCriteriaSet In pCriteriaSets
            ClearTableAliases()
            vAttrList = ""
            If mvAppealMailing Then vAttrList = pSegment.SegmentSequence & ", "
            vAttrList = vAttrList & pSegment.SelectionSet & ", 1, "
            'Create the sql statement for this criteria set
            BuildStatement(IntegerValue(vCriteriaSet))
            If pSegment.Score.Length > 0 Then
              mvTableNumber = mvTableNumber + 1
              mvTableList = mvTableList & ", contact_scores table" & mvTableNumber
              If Not mvTableAliases.Exists("table" & mvTableNumber) Then mvTableAliases.Add("contact_scores", "table" & mvTableNumber)
              mvWhere = mvWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = table" & mvTableNumber & "." & mvContactAttribute
              mvWhere = mvWhere & " AND table" & mvTableNumber & ".score = '" & pSegment.Score & "'"
            End If
            If Not mvAppealMailing And (mvMailingType <> MailingTypes.mtyMemberFulfilment And mvMailingType <> MailingTypes.mtyNonMemberFulfilment And mvMailingType <> MailingTypes.mtySelectionTester) Then 'Selection Manager Mailing, but not Member Fulfilment, Non-member Fulfilment or Selection Tester
              If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
                vMailingAccess = mvEnv.GetConfig("ownership_mailings")
                If Len(vMailingAccess) > 0 Then
                  'Need to join thru the contacts table to get the ownership group
                  'vPos = InStr(mvTableList, "contacts table")
                  'BR15112: It looks like mvTableList will always start with an empty space, but just in case check if there is no space 
                  'and the first word is "contacts table". This will also make sure the system does not read selected_contacts table as contacts table
                  If mvTableList.StartsWith("contacts table") Then
                    vPos = 1
                  Else
                    vPos = InStr(mvTableList, " contacts table")
                    If vPos > 0 Then vPos += 1 'add one for empty space
                  End If
                  If vPos > 0 Then
                    vCTable = "table" & Val(Mid(mvTableList, vPos + 14))
                  Else
                    mvTableNumber = mvTableNumber + 1
                    mvTableList = mvTableList & ", contacts table" & mvTableNumber
                    If Not mvTableAliases.Exists("table" & mvTableNumber) Then mvTableAliases.Add("contacts", "table" & mvTableNumber)
                    mvWhere = mvWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = table" & mvTableNumber & "." & mvContactAttribute
                    vCTable = "table" & mvTableNumber
                  End If
                  mvTableNumber = mvTableNumber + 1
                  mvTableList = mvTableList & ", ownership_group_users table" & mvTableNumber
                  If Not mvTableAliases.Exists("table" & mvTableNumber) Then mvTableAliases.Add("ownership_group_users", "table" & mvTableNumber)
                  mvWhere = mvWhere & " AND " & vCTable & ".ownership_group = table" & mvTableNumber & ".ownership_group"
                  mvWhere = mvWhere & " AND table" & mvTableNumber & ".logname = '" & mvEnv.User.Logname & "'"
                  mvWhere = mvWhere & " AND table" & mvTableNumber & ".ownership_access_level >= '" & vMailingAccess & "'"
                  mvWhere = mvWhere & " AND table" & mvTableNumber & ".valid_from" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (TodaysDate()))
                  mvWhere = mvWhere & " AND (( table" & mvTableNumber & ".valid_to" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (TodaysDate())) & " )"
                  mvWhere = mvWhere & " OR (table" & mvTableNumber & ".valid_to IS NULL))"
                End If
              Else
                mvTableNumber = mvTableNumber + 1
                mvTableList = mvTableList & ", contact_users table" & mvTableNumber
                If Not mvTableAliases.Exists("table" & mvTableNumber) Then mvTableAliases.Add("contact_users", "table" & mvTableNumber)
                mvWhere = mvWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = table" & mvTableNumber & "." & mvContactAttribute
                mvWhere = mvWhere & " AND table" & mvTableNumber & ".department = '" & mvEnv.User.Department & "'"
              End If
            End If
            vAddrTable = LinkToAddress()
            'If mvMasterAttribute <> mvContactAttribute Then vAttrList = vAttrList & mvMasterTable & "." & mvMasterAttribute & ", "
            If mvLinkToContacts And Not mvMasterHasAddress Then
              If mvMasterAttribute <> mvContactAttribute Then vAttrList = vAttrList & mvMasterTableAlias & "." & mvMasterAttribute & ", "
            Else
              If mvMasterAttribute <> mvContactAttribute Then
                vAttrList = vAttrList & TableContainsMaster(mvContactTableAlias) & "." & mvMasterAttribute & ", "
              End If
            End If
            vAttrList = vAttrList & mvContactTableAlias & "." & mvContactAttribute & ", "
            vAttrList = vAttrList & vAddrTable & "." & mvAddressAttribute 'once to populate the address_number attribute...
            vAttrList = vAttrList & ", " & vAddrTable & "." & mvAddressAttribute '...and once to populate the address_number_2 attribute
            If mvMailingType = MailingTypes.mtyIrishGiftAid Then vAttrList = vAttrList & ", " & mvMasterTableAlias & ".performance"
            vSQL = "INSERT INTO " & pTableName & "(" & vDestAttrs & ") SELECT " & vAttrList & " FROM " & mvTableList & mvWhere
            mvConn.LogMailSQL(vSQL)
            vCount = vCount + mvConn.ExecuteSQL(vSQL)
            'Remove temporary criteria set
            vWhereFields = New CDBFields
            vWhereFields.Add("criteria_set", CDBField.FieldTypes.cftLong, vCriteriaSet)
            mvConn.DeleteRecords("criteria_set_details", vWhereFields)
          Next vCriteriaSet
        End If
      End If
      mvDestinationAttributes = vDestAttrs
      mvSelectionAttributes = vAttrList
    End Sub

    Public Sub ProcessStep(ByRef pSelectionTable As String, ByRef pDestAttrs As String, ByRef pAttrList As String, ByRef pSegment As Segment, ByRef pStep As SelectionStep, ByRef pAppealMailing As Boolean, Optional ByRef pLogMail As Boolean = False)
      Dim vSQL As String
      Dim vAddrUsage As String


      Select Case pStep.SelectAction
        Case "S"
          vSQL = "INSERT INTO " & pSelectionTable & "(" & pDestAttrs & ") SELECT DISTINCT " & Replace(pAttrList, "x." & mvAddressAttribute, "MIN(" & "x." & mvAddressAttribute & ")") & " FROM " & pStep.ViewName & " x "
          If pSegment.Score.Length > 0 Then vSQL = vSQL & ", contact_scores cs "
          vSQL = vSQL & StepOwnershipRestriction()
          If pSegment.Score.Length > 0 Then
            vSQL = vSQL & "x." & mvContactAttribute & " = cs." & mvContactAttribute & " AND cs.score = '" & pSegment.Score & "' AND "
          End If
          vSQL = vSQL & pStep.GetAdjustedFilterSQL()
          vSQL = vSQL & " AND x.contact_number NOT IN ( SELECT contact_number FROM " & pSelectionTable & " WHERE selection_set = " & pSegment.SelectionSet
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          vSQL = vSQL & ") GROUP BY x." & mvContactAttribute
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)

        Case "R"
          vSQL = "DELETE FROM " & pSelectionTable & " WHERE selection_set = " & pSegment.SelectionSet
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          vSQL = vSQL & " AND contact_number IN ( SELECT DISTINCT x.contact_number FROM " & pSelectionTable & " x, " & pStep.ViewName & " vn "
          vSQL = vSQL & "WHERE x.selection_set = " & pSegment.SelectionSet
          If pAppealMailing Then vSQL = vSQL & " AND x.segment_sequence = " & pSegment.SegmentSequence
          vSQL = vSQL & " AND x.contact_number = vn.contact_number AND " & pStep.GetAdjustedFilterSQL() & ")"
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)

        Case "P"
          'Mark records to keep as revision 2
          vSQL = "UPDATE " & pSelectionTable & " SET revision = 2 WHERE selection_set = " & pSegment.SelectionSet
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          vSQL = vSQL & " AND contact_number IN ( SELECT DISTINCT x.contact_number FROM " & pSelectionTable & " x, " & pStep.ViewName & " vn "
          vSQL = vSQL & " WHERE x.selection_set = " & pSegment.SelectionSet
          vSQL = vSQL & " AND x.contact_number = vn.contact_number AND " & pStep.GetAdjustedFilterSQL() & ")"
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)
          'Delete the revision 1 records
          vSQL = "DELETE FROM " & pSelectionTable & " WHERE selection_set = " & pSegment.SelectionSet & " AND revision = 1"
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)
          'Update revision 2 to revision 1
          vSQL = "UPDATE " & pSelectionTable & " SET revision = 1 WHERE selection_set = " & pSegment.SelectionSet
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)

        Case "D"
          pAttrList = Replace(pAttrList, "x.address_number", "c.address_number")
          vSQL = "INSERT INTO " & pSelectionTable & "(" & pDestAttrs & ") SELECT DISTINCT " & Replace(pAttrList, ",1,", ",2,") & " FROM " & pSelectionTable & " x, contacts c, contact_addresses ca"
          vSQL = vSQL & " WHERE selection_set = " & pSegment.SelectionSet & " AND revision = 1"
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          vSQL = vSQL & " AND x.contact_number = c.contact_number AND x.contact_number = ca.contact_number AND c.address_number = ca.address_number AND ca.historical = 'N'"
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)
          vSQL = "DELETE FROM " & pSelectionTable & " WHERE selection_set = " & pSegment.SelectionSet & " AND revision = 1"
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)
          vSQL = "UPDATE " & pSelectionTable & " SET revision = 1 WHERE selection_set = " & pSegment.SelectionSet & " AND revision = 2"
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)

        Case "U"
          vAddrUsage = Left(Mid(pStep.FilterSql, InStr(pStep.FilterSql, "'") + 1), Len(Mid(pStep.FilterSql, InStr(pStep.FilterSql, "'") + 1)) - 1)
          vSQL = "UPDATE " & pSelectionTable & " SET address_number = (SELECT max(ca.address_number) from contact_addresses ca, contact_address_usages cau WHERE ca.contact_number = " & pSelectionTable & ".contact_number AND ca.historical = 'N' AND ca.contact_number = cau.contact_number and ca.address_number = cau.address_number and address_usage = '" & vAddrUsage & "')" 'update address_number to the highet address number with the address usage for this contact
          vSQL = vSQL & "WHERE contact_number IN (SELECT ca.contact_number from contact_addresses ca, contact_address_usages cau WHERE ca.contact_number = " & pSelectionTable & ".contact_number AND ca.historical = 'N' AND ca.contact_number = cau.contact_number and ca.address_number = cau.address_number and address_usage = '" & vAddrUsage & "') " 'Only update the address for contacts that actually have the specified address usage and leave the others as they were
          vSQL = vSQL & "AND contact_number NOT IN (SELECT ca1.contact_number FROM contact_addresses ca1, contact_address_usages cau1 WHERE ca1.contact_number = " & pSelectionTable & ".contact_number AND ca1.address_number =  " & pSelectionTable & ".address_number AND historical = 'N' AND ca1.contact_number = cau1.contact_number AND ca1.address_number = cau1.address_number AND cau1.address_usage = '" & vAddrUsage & "')" 'also exclude contacts for which we already have a address with the specified address usage, to avoid over writing it with a different one.
          If pAppealMailing Then vSQL = vSQL & " AND segment_sequence = " & pSegment.SegmentSequence
          If pLogMail Then mvConn.LogMailSQL(vSQL)
          mvConn.ExecuteSQL(vSQL)

        Case "N"
          'List manager Random Data Sample
          LMSelectRandomContacts(pSegment.SelectionSet, pSelectionTable, pStep.FilterSql)

      End Select

      Dim vIndexes As New CDBIndexes
      vIndexes.Init(mvConn, pSelectionTable)
      If vIndexes.Count = 0 Then vIndexes.CreateIfMissing(mvConn, False, {"contact_number"})

    End Sub

    Public Sub ProcessEbuStep(ByRef pSelectionTable As String, ByRef pDestAttrs As String, ByRef pAttrList As String, ByRef pSegment As Segment, ByRef pStep As SelectionStep)
      Dim vSQL As String
      Select Case pStep.SelectAction
        Case "S"
          vSQL = "INSERT INTO " & pSelectionTable & "(" & pDestAttrs & ") SELECT DISTINCT " & pAttrList & " FROM " & pStep.ViewName & " x "
          vSQL = vSQL & "WHERE " & pStep.GetAdjustedFilterSQL()
          vSQL = vSQL & " AND x.exam_booking_unit_id NOT IN ( SELECT exam_booking_unit_id FROM " & pSelectionTable & " WHERE selection_set = " & pSegment.SelectionSet
          vSQL = vSQL & ") GROUP BY x.exam_booking_unit_id"
          mvConn.ExecuteSQL(vSQL)
#If DEBUG Then
        Case Else
          Debugger.Break()
#End If
      End Select

      Dim vIndexes As New CDBIndexes
      vIndexes.Init(mvConn, pSelectionTable)
      If vIndexes.Count = 0 Then
        vIndexes.CreateIfMissing(mvConn, False, {"exam_booking_unit_id"})
      End If

    End Sub

    Public Function StepOwnershipRestriction() As String
      Dim vWhere As String
      Dim vMailingAccess As String

      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        vMailingAccess = mvEnv.GetConfig("ownership_list_manager")
        If Len(vMailingAccess) > 0 Then
          vWhere = ", ownership_group_users ogu WHERE x.ownership_group = ogu.ownership_group"
          vWhere = vWhere & " AND ogu.logname = '" & mvEnv.User.Logname & "'"
          vWhere = vWhere & " AND ogu.ownership_access_level >= '" & vMailingAccess & "'"
          vWhere = vWhere & " AND ogu.valid_from" & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, (TodaysDate()))
          vWhere = vWhere & " AND (( ogu.valid_to" & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (TodaysDate())) & " )"
          vWhere = vWhere & " OR (ogu.valid_to IS NULL)) AND "
          StepOwnershipRestriction = vWhere
        Else
          StepOwnershipRestriction = "WHERE "
        End If
      Else
        StepOwnershipRestriction = ", contact_users cu WHERE x.contact_number = cu.contact_number AND cu.department = '" & mvEnv.User.Department & "' AND "
      End If
    End Function

    Private Sub LMSelectRandomContacts(ByRef pSelectionSet As Integer, ByVal pTable As String, ByVal pRequiredCount As String)

      Dim vString As String() = pRequiredCount.Split(" "c)
      Dim vRequiredCount As Integer = CInt(vString(2))
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftInteger, pSelectionSet)
      Dim vActualCount As Integer = mvConn.GetCount(pTable, vWhereFields)

      'if the number passed through is a percentage, then calculate number of records required
      If vString(0) = "SamplePercentage" Then
        vRequiredCount = CInt(CDbl(vActualCount) / 100 * (vRequiredCount))
      End If

      'if our sample is more records than we have in our selection set then don't do anything
      If vActualCount <= vRequiredCount Then Exit Sub

      Dim vExclude As Boolean
      If vRequiredCount > (vActualCount - vRequiredCount) Then
        'you're including more than you're excluding so only mark those that should be excluded
        vExclude = True
        vRequiredCount = vActualCount - vRequiredCount
      End If

      'generate random numbers and put them in a list
      Dim vList As New SortedList(Of Integer, Integer)
      Dim vRand As Integer
      While vList.Count < vRequiredCount
        vRand = GetRandomInt(0, vActualCount)
        If Not vList.ContainsKey(vRand) Then vList.Add(vRand, vRand)
      End While

      'go through the record set and get the contact numbers for each of the rows that needs updating in the recordset
      Dim vRecordSet As CDBRecordSet
      vRecordSet = mvConn.GetRecordSet("SELECT contact_number FROM " & pTable & " WHERE selection_set = " & pSelectionSet, CDBConnection.RecordSetOptions.NoDataTable)
      Dim vCount As Integer = 0
      With vRecordSet
        While .Fetch
          If vList.ContainsKey(vCount) Then
            vList(vCount) = vRecordSet.Fields("contact_number").IntegerValue
          End If
          vCount += 1
        End While
        .CloseRecordSet()
      End With

      'make sure marker column is null for all rows in case we're doing a sample of a sample
      Dim vUpdateFields As New CDBFields
      vUpdateFields.Add("marker", CDBField.FieldTypes.cftInteger)
      mvConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
      'now set marker to 1 for update to records
      vUpdateFields.Item("marker").Value = "1"

      mvConn.LogMailSQL("Starting the Selection of Random Contacts")

      'update the marker column for the ones that need to be marked in the database
      vWhereFields = New CDBFields
      vWhereFields.Add("contact_number")
      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSelectionSet)
      vWhereFields.Add("marker", 1, CDBField.FieldWhereOperators.fwoNullOrNotEqual)
      For Each vItem As KeyValuePair(Of Integer, Integer) In vList
        vWhereFields(1).Value = vItem.Value.ToString
        mvEnv.Connection.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
      Next

      'delete the record that need deleting using the marker
      vWhereFields = New CDBFields
      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSelectionSet)
      If vExclude Then
        vWhereFields.Add("marker", CDBField.FieldTypes.cftInteger, 1)
      Else
        vWhereFields.Add("marker", CDBField.FieldTypes.cftInteger)
      End If
      mvConn.DeleteRecords(pTable, vWhereFields)

      mvConn.LogMailSQL("List Manager Random Data Sample Completed")
    End Sub

    Private Function GetRandomInt(ByVal pMin As Integer, ByVal pMax As Integer) As Integer
      Static Generator As System.Random = New System.Random()
      Return Generator.Next(pMin, pMax)
    End Function

    Private Sub SelectRandomContacts(ByRef pSegment As Segment, ByVal pTable As String)
      Dim vCount As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vDivider As Integer
      Dim vUpdateFields As CDBFields
      Dim vWhereFields As CDBFields
      Dim vQuit As Boolean
      Dim vExclude As Boolean
      Dim vRecCount As Integer
      Dim vUpdate As Boolean
      Dim vStartRecord As Integer

      mvConn.LogMailSQL("Starting the Selection of Random Contacts")

      If pSegment.RequiredCount > (pSegment.ActualCount - pSegment.RequiredCount) Then
        'you're including more than you're excluding so only mark those that should be excluded
        vExclude = True
        vRecCount = pSegment.ActualCount
      Else
        'you're including less than you're excluding so only mark those that should be included
        vRecCount = 0
      End If

      vUpdateFields = New CDBFields
      vUpdateFields.Add("marker", CDBField.FieldTypes.cftCharacter, "X")

      vDivider = CInt(System.Math.Round(System.Math.Abs(pSegment.ActualCount / pSegment.RequiredCount)) - 1)
      If vDivider < 2 Then vDivider = 2 'do this because you don't want to divide by zero or one
      vStartRecord = CInt(Int((vDivider * Rnd()) + 1))

      vRecordSet = mvConn.GetRecordSet("SELECT selection_set," & mvMasterAttribute & " FROM " & pTable & " WHERE selection_set = " & pSegment.SelectionSet)
      With vRecordSet
        While .Fetch() = True And Not vQuit
          vWhereFields = New CDBFields
          vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, .Fields(1).IntegerValue)
          vWhereFields.Add("revision", CDBField.FieldTypes.cftInteger, "1")
          vWhereFields.Add(mvMasterAttribute, CDBField.FieldTypes.cftLong, .Fields(2).IntegerValue)
          vUpdate = False
          vCount = vCount + 1
          If vCount = vStartRecord Then 'always include this record
            vCount = 0 'restart numbering sequence
            vStartRecord = pSegment.ActualCount + 1 'ensure the starting point is not reached again
            If Not vExclude Then
              vUpdate = True
              vRecCount = vRecCount + 1
            End If
          ElseIf vCount Mod vDivider > 0 Then
            If vExclude Then
              vUpdate = True
              vRecCount = vRecCount - 1
            End If
          Else
            If Not vExclude Then
              vUpdate = True
              vRecCount = vRecCount + 1
            End If
          End If
          If vUpdate Then mvConn.UpdateRecords(pTable, vUpdateFields, vWhereFields)
          If vRecCount = pSegment.RequiredCount Then vQuit = True
        End While
        .CloseRecordSet()
      End With

      vWhereFields = New CDBFields
      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSegment.SelectionSet)
      If vExclude Then
        vWhereFields.Add("marker", CDBField.FieldTypes.cftCharacter, "X")
      Else
        vWhereFields.Add("marker", CDBField.FieldTypes.cftCharacter)
      End If
      mvConn.DeleteRecords(pTable, vWhereFields)

      pSegment.ActualCount = pSegment.RequiredCount

      mvConn.LogMailSQL("Random Contact Selection Completed")
    End Sub
    Private Sub BuildAttrCriteria(ByRef pCC As CriteriaContext, ByRef pSelectTable As String)

      If pCC.MainValue <> "" Then
        ProcessMainValue(pCC, (pCC.MainValue), pSelectTable)
        If pCC.SubValue <> "" Then
          mvSingleWhere = mvSingleWhere & " AND"
          ProcessSubsidiaryValue(pCC, (pCC.MainValue), (pCC.SubValue), pSelectTable)
        End If
        If pCC.Period <> "" Then
          mvSingleWhere = mvSingleWhere & " AND "
          ProcessDates(pCC, pSelectTable)
        End If
      Else
        ProcessDates(pCC, pSelectTable)
      End If
    End Sub
    Private Sub BuildExcludeBody(ByRef pCC As CriteriaContext)
      Dim vSelectTable As String
      Dim vTableNumber As Integer

      vTableNumber = mvTableNumber
      If mvTableNumber = 1 Then
        vSelectTable = "table1"
        If pCC.Contacts Then
          If pCC.SearchArea = "role" Then
            mvTableList = mvPositionTable & " table1"
            If Not mvTableAliases.Exists("table1") Then mvTableAliases.Add(mvPositionTable, "table1")
            mvLastOTable = vSelectTable
            If mvCAddressLink = "" Then
              mvCAddressLink = vSelectTable
            End If
          Else
            mvTableList = mvMasterTable & " table1"
            mvMasterTableAlias = "table1"
            If Not mvTableAliases.Exists("table1") Then mvTableAliases.Add(mvMasterTable, "table1")
          End If
          mvWhere = " WHERE " & TableContainsMaster(vSelectTable) & "." & mvMasterAttribute & " NOT IN (SELECT table" & CStr(mvTableNumber + 1) & "." & mvMasterAttribute & " FROM "
          mvLastCTable = vSelectTable
        Else
          mvTableList = mvPositionTable & " table1"
          If Not mvTableAliases.Exists("table1") Then mvTableAliases.Add(mvPositionTable, "table1")
          mvWhere = " WHERE (" & vSelectTable & "." & mvOrgAttribute & " NOT IN (SELECT table" & CStr(mvTableNumber + 1) & "." & mvOrgAttribute & " FROM "
          mvLastCTable = vSelectTable
          mvLastOTable = vSelectTable
          If mvCAddressLink = "" Then
            mvCAddressLink = vSelectTable
          End If
        End If
        mvContactTableAlias = "table1"
        mvTableNumber = mvTableNumber + 1
      Else
        If pCC.Contacts Then
          If mvLastCTable = "" Then
            ' must have joined to organisations - now link though positions table
            vSelectTable = "table" & CStr(mvTableNumber)
            mvWhere = mvWhere & " AND " & mvLastOTable & "." & mvOrgAttribute & " = " & vSelectTable & ".organisation_number AND " & mvConn.DBSpecialCol(vSelectTable, "current") & " = 'Y' AND " & vSelectTable & "." & "mail = 'Y'"
            mvWhere = mvWhere & " AND " & vSelectTable & "." & mvContactAttribute & " NOT IN (SELECT table" & CStr(mvTableNumber + 1) & "." & mvContactAttribute & " FROM "
            mvTableList = mvTableList & ", " & mvPositionTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvPositionTable, vSelectTable)
            mvLastCTable = vSelectTable
            If mvCAddressLink = "" Then
              mvCAddressLink = vSelectTable
            End If
            mvTableNumber = mvTableNumber + 1
          Else
            vSelectTable = mvLastCTable
            mvWhere = mvWhere & " AND " & TableContainsMaster(vSelectTable) & "." & mvMasterAttribute & " NOT IN (SELECT table" & CStr(mvTableNumber) & "." & mvMasterAttribute & " FROM "
          End If
        Else
          If mvLastOTable = "" Then
            ' must have joined to contacts - now link though positions table
            vSelectTable = "table" & CStr(mvTableNumber)
            mvWhere = mvWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = " & vSelectTable & "." & mvContactAttribute & " AND " & mvConn.DBSpecialCol(vSelectTable, "current") & " = 'Y' AND " & vSelectTable & "." & "mail = 'Y' AND " & vSelectTable & "." & mvOrgAttribute & " NOT IN (SELECT table" & CStr(mvTableNumber + 1) & "." & mvOrgAttribute & " FROM "
            'Next line introduced to fix a bug in the 4GL 28/7/95
            mvLastCTable = vSelectTable
            mvTableList = mvTableList & ", " & mvPositionTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvPositionTable, vSelectTable)
            mvLastOTable = vSelectTable
            If mvCAddressLink = "" Then
              mvCAddressLink = vSelectTable
            End If
            mvTableNumber = mvTableNumber + 1
          Else
            vSelectTable = mvLastOTable
            mvWhere = mvWhere & " AND " & vSelectTable & "." & mvOrgAttribute & " NOT IN (SELECT table" & CStr(mvTableNumber) & "." & mvOrgAttribute & " FROM "
          End If
        End If
      End If
      vSelectTable = "table" & CStr(mvTableNumber)
      mvSingleWhere = " WHERE ("
      If pCC.Contacts Then
        mvSingleWhere = mvSingleWhere & TableContainsMaster(vSelectTable) & "." & mvMasterAttribute & " = " & TableContainsMaster(mvLastCTable) & "." & mvMasterAttribute
      Else
        mvSingleWhere = mvSingleWhere & vSelectTable & "." & mvOrgAttribute & " = " & mvLastOTable & "." & mvOrgAttribute
      End If
      If pCC.Special Then
        mvSingleWhere = mvSingleWhere & " AND "
        SpecialLater(pCC)
        vSelectTable = "table" & CStr(mvTableNumber)
      Else
        mvSingleTableList = pCC.TableName & " " & vSelectTable
        If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add((pCC.TableName), vSelectTable)
        If pCC.SearchArea = "role" Then
          mvTableNumber = mvTableNumber + 1
          mvSingleTableList = mvSingleTableList & ", " & mvPositionTable & " " & ("table" & CStr(mvTableNumber))
          If Not mvTableAliases.Exists("table" & CStr(mvTableNumber)) Then mvTableAliases.Add(mvPositionTable, "table" & CStr(mvTableNumber))
          mvSingleWhere = mvSingleWhere & " AND " & vSelectTable & "." & mvContactAttribute & " = " & ("table" & CStr(mvTableNumber)) & "." & mvContactAttribute & " AND " & vSelectTable & "." & mvOrgAttribute & " = " & ("table" & CStr(mvTableNumber)) & "." & mvOrgAttribute & " AND " & mvConn.DBSpecialCol("table" & CStr(mvTableNumber), "current") & " = 'Y' AND " & ("table" & CStr(mvTableNumber)) & "." & "mail = 'Y'"
        End If
        '    mvSingleTableList = pCC.TableName & " " & vSelectTable
      End If
      mvSingleWhere = mvSingleWhere & " AND "
      BuildAttrCriteria(pCC, vSelectTable)
      mvSingleWhere = mvSingleWhere & ") "
      mvWhere = mvWhere & mvSingleTableList
      mvSingleTableList = " "
      mvSingleWhere = mvSingleWhere & ")"
      If vTableNumber = 1 Then
        If pCC.Contacts = False Then mvSingleWhere = mvSingleWhere & ") AND " & mvConn.DBSpecialCol("table1", "current") & " = 'Y' AND table1.mail = 'Y'"
      End If
    End Sub
    Private Sub BuildIncludeBody(ByRef pCC As CriteriaContext, ByRef pLookForward As Boolean)
      Dim vSelectTable As String

      If mvTableNumber = 1 Then
        mvSingleTableList = pCC.TableName & " table1"
        vSelectTable = "table1"
        If pCC.TableName = mvMasterTable Then
          mvMasterTableAlias = vSelectTable
          mvContactTableAlias = vSelectTable
        End If
        If pCC.TableName = mvOrgTable Then mvOrgTableAlias = vSelectTable
        mvSingleWhere = " WHERE ("
        StoreAddressLinks(pCC, vSelectTable)
        If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add((pCC.TableName), vSelectTable)
      Else
        If mvMasterTableAlias.Length = 0 AndAlso pCC.TableName = mvMasterTable Then
          vSelectTable = "table" & CStr(mvTableNumber)
          mvMasterTableAlias = vSelectTable
          mvContactTableAlias = vSelectTable
        End If
        If pCC.SearchArea = "role" Then
          If mvLastCTable = "" Then pCC.Contacts = False 'fool the system
        End If

        If pCC.Contacts Then
          If mvLastCTable = "" Then
            'must have joined to organisations - now link though positions table
            vSelectTable = "table" & CStr(mvTableNumber)
            mvSingleWhere = mvSingleWhere & " AND " & mvLastOTable & "." & mvOrgAttribute & " = " & vSelectTable & ".organisation_number AND " & mvConn.DBSpecialCol(vSelectTable, "current") & " = 'Y' AND " & vSelectTable & "." & "mail = 'Y' AND " & vSelectTable & "." & mvContactAttribute & " = "
            mvSingleTableList = mvSingleTableList & ", " & mvPositionTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvPositionTable, vSelectTable)
            If mvCAddressLink = "" Then mvCAddressLink = vSelectTable
            mvTableNumber = mvTableNumber + 1
          Else
            If pCC.SearchArea = "role" Then
              If mvLastOTable <> "" Then
                vSelectTable = "table" & CStr(mvTableNumber)
                mvSingleWhere = mvSingleWhere & " AND " & mvLastOTable & "." & mvOrgAttribute & " = " & vSelectTable & "." & mvOrgAttribute
              End If
            End If
            vSelectTable = mvLastCTable
            mvSingleWhere = mvSingleWhere & " AND " & TableContainsMaster(vSelectTable) & "." & mvMasterAttribute & " = "
          End If
        Else
          If mvLastOTable = "" Then
            'must have joined to contacts - now link though positions table
            vSelectTable = "table" & CStr(mvTableNumber)
            mvSingleWhere = mvSingleWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = " & vSelectTable & "." & mvContactAttribute & " AND " & mvConn.DBSpecialCol(vSelectTable, "current") & " = 'Y' AND " & vSelectTable & "." & "mail = 'Y' AND " & vSelectTable & "." & mvOrgAttribute & " = "
            mvSingleTableList = mvSingleTableList & ", " & mvPositionTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvPositionTable, vSelectTable)
            If mvCAddressLink = "" Then mvCAddressLink = vSelectTable
            mvTableNumber = mvTableNumber + 1
          Else
            If pCC.SearchArea = "role" Then
              If mvLastCTable <> "" Then
                vSelectTable = "table" & CStr(mvTableNumber)
                mvSingleWhere = mvSingleWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = " & vSelectTable & "." & mvContactAttribute
              End If
            End If
            vSelectTable = mvLastOTable
            mvSingleWhere = mvSingleWhere & " AND " & vSelectTable & "." & mvOrgAttribute & " = "
          End If
        End If
        vSelectTable = "table" & CStr(mvTableNumber)
        If mvMailingType = MailingTypes.mtyIrishGiftAid Then
          'If the MasterTable was not the first table then mvMastertableAlias will not have been set
          'For Irish Gift Aid this must be set but don't want to set it globally as I don't know if this will break other mailings
          If Len(mvMasterTableAlias) = 0 And (pCC.TableName = mvMasterTable) Then
            mvMasterTableAlias = vSelectTable
          End If
        End If
        If pCC.Contacts Then
          mvLastCTable = vSelectTable
          mvLastCSearchArea = pCC.SearchArea
          mvSingleWhere = mvSingleWhere & TableContainsMaster(vSelectTable) & "." & mvMasterAttribute
        Else
          mvLastOTable = vSelectTable
          mvLastOSearchArea = pCC.SearchArea
          mvSingleWhere = mvSingleWhere & vSelectTable & "." & mvOrgAttribute
        End If
        If pCC.Special Then
          mvSingleWhere = mvSingleWhere & " AND "
          SpecialLater(pCC)
          vSelectTable = "table" & CStr(mvTableNumber)
        Else
          mvSingleTableList = mvSingleTableList & ", " & pCC.TableName & " " & vSelectTable
          If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add((pCC.TableName), vSelectTable)
          If pCC.TableName = mvMasterTable Then mvContactTableAlias = vSelectTable
          If pCC.TableName = mvOrgTable Then mvOrgTableAlias = vSelectTable
          StoreAddressLinks(pCC, vSelectTable)
        End If
        mvSingleWhere = mvSingleWhere & " AND ("
      End If

      BuildAttrCriteria(pCC, vSelectTable)
      If pLookForward Then ProcessSameTable(pCC, vSelectTable, (pCC.TableName), (pCC.SearchArea), (pCC.MainAttribute), (pCC.MainValue))
      mvSingleWhere = mvSingleWhere & ") "
      If pCC.SearchArea = "role" Then
        If mvLastOTable = "" Or mvLastCTable = "" Then
          If pCC.Special Then SpecialFirst(pCC)
          mvLastCTable = vSelectTable
          mvTableNumber = mvTableNumber + 1
          vSelectTable = "table" & CStr(mvTableNumber)
          mvSingleTableList = mvSingleTableList & ", " & mvPositionTable & " " & vSelectTable
          If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvPositionTable, vSelectTable)
          If mvCAddressLink = "" Then mvCAddressLink = vSelectTable
          mvSingleWhere = mvSingleWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = " & vSelectTable & "." & mvContactAttribute & " AND " & mvLastCTable & "." & mvOrgAttribute & " = " & vSelectTable & "." & mvOrgAttribute & " AND " & mvConn.DBSpecialCol(vSelectTable, "current") & " = 'Y' AND " & vSelectTable & "." & "mail = 'Y'"
        End If
        mvLastCTable = vSelectTable
        mvLastOTable = vSelectTable
      End If
      If mvTableNumber = 1 Then
        If pCC.Special Then SpecialFirst(pCC)
        'record last table so I know what to join to - next time around
        If pCC.Contacts Then
          mvLastCTable = "table" & CStr(mvTableNumber)
          mvLastCSearchArea = pCC.SearchArea
        Else
          mvLastOTable = "table" & CStr(mvTableNumber)
          mvLastOSearchArea = pCC.SearchArea
        End If
      Else
        StoreAddressLinks(pCC, vSelectTable)
      End If
    End Sub
    Private Function BuildMainList(ByRef pCC As CriteriaContext, ByRef pMainValue As String, ByRef pDelimiter As String) As Boolean
      Dim vTable As String
      Dim vMainAttr As String
      Dim vMainAttrType As String
      Dim vValAttr As String
      Dim vList As String
      Dim vToken As String
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      vMainAttr = pCC.MainAttribute
      vMainAttrType = pCC.MainDataType
      vValAttr = pCC.MainValidationAttribute
      vTable = pCC.ValidationTable

      mvPosition = 1
      mvToken1 = ""
      mvToken2 = ""
      vList = "'"

      Do
        If mvEnv.Connection.IsUnicodeField(pCC.MainAttribute) Then
          vToken = NextToken(pMainValue, TokenActionTypes.tatSubstitute, True)
        Else
          vToken = NextToken(pMainValue, TokenActionTypes.tatSubstitute)
        End If
        If vToken <> "" Then
          If mvToken2 <> "" Then
            vSQL = "SELECT * FROM " & vTable & " WHERE " & vTable & "." & vValAttr & " BETWEEN " & pDelimiter & mvToken1 & pDelimiter & " AND " & pDelimiter & mvToken2 & pDelimiter
          Else
            If vValAttr <> vMainAttr Then
              vSQL = "SELECT * FROM " & vTable & " WHERE " & vTable & "." & vValAttr & " = " & pDelimiter & mvToken1 & pDelimiter
            Else
              vSQL = ""
            End If
          End If
          If vSQL.Length > 0 Then
            vRecordSet = mvConn.GetRecordSet(vSQL)
            While vRecordSet.Fetch() = True
              If vRecordSet.Fields.Exists(vMainAttr) Then 'vMainAttr may not exist in the validation table due to difference in attribute names
                If vMainAttrType <> "C" Then
                  If vList = "'" Then
                    vList = vRecordSet.Fields(vMainAttr).Value
                  Else
                    vList = vList & "," & vRecordSet.Fields(vMainAttr).Value
                  End If
                Else
                  If vList = "'" Then
                    vList = vList & vRecordSet.Fields(vMainAttr).Value & "'"
                  Else
                    vList = vList & ",'" & vRecordSet.Fields(vMainAttr).Value & "'"
                  End If
                End If
              End If
            End While
            vRecordSet.CloseRecordSet()
          End If
        End If
      Loop While vToken <> ""

      If vList <> "'" Then
        mvSingleWhere = mvSingleWhere & vList
        BuildMainList = True
      End If
    End Function
    Private Sub BuildStatement(ByRef pCriteriaSet As Integer)
      Dim vCC As CriteriaContext

      mvTableNumber = 0
      mvLastCTable = ""
      mvLastOTable = ""
      mvCGeographic = ""
      mvOGeographic = ""
      mvCAddressLink = ""
      mvOAddressLink = ""
      mvContactTableAlias = ""
      mvOrgTableAlias = ""
      mvLastCSearchArea = ""
      mvLastOSearchArea = ""
      mvTableList = " "
      mvWhere = " "

      GetCriteriaDetails(pCriteriaSet, False)
      For Each vCC In mvCriteriaContexts
        If Not vCC.Processed Then
          mvTableNumber = mvTableNumber + 1
          BuildStatementBody(vCC, True)
        End If
      Next vCC
      LinkToPositions()
      LinkToMaster()
      LinkToContacts()
    End Sub
    Private Sub BuildStatementBody(ByRef pCC As CriteriaContext, ByRef pLookForward As Boolean)

      mvSingleTableList = " "
      mvSingleWhere = " "
      If pCC.Include Then
        BuildIncludeBody(pCC, pLookForward)
      Else
        BuildExcludeBody(pCC)
      End If
      mvTableList = mvTableList & LTrim(mvSingleTableList)

      mvWhere = mvWhere & mvSingleWhere
    End Sub
    Private Function BuildSubsidiaryList(ByRef pCC As CriteriaContext, ByRef pMainValue As String, ByRef pSubsidiaryValue As String) As Boolean
      Dim vTable As String
      Dim vMainPosition As Integer
      Dim vMainAttr As String
      Dim vSubsidiaryAttr As String
      Dim vSubAttrType As String
      Dim vMainToken1 As String
      Dim vMainToken2 As String
      Dim vList As String
      Dim vToken As String
      Dim vToken2 As String
      Dim vSQL As String
      Dim vRecordSet As CDBRecordSet

      vTable = pCC.ValidationTable
      vMainAttr = pCC.MainValidationAttribute
      vSubsidiaryAttr = pCC.SubValidationAttribute
      vSubAttrType = pCC.SubDataType

      vMainPosition = 1
      mvToken1 = ""
      mvToken2 = ""
      mvPosition = 1
      vList = "'"
      Do
        If mvEnv.Connection.IsUnicodeField(pCC.SubAttribute) Then
          vToken = NextToken(pMainValue, TokenActionTypes.tatSubstitute, True)
        Else
          vToken = NextToken(pMainValue, TokenActionTypes.tatSubstitute)
        End If
        If vToken <> "" Then
          vMainToken1 = mvToken1
          vMainToken2 = mvToken2
          vMainPosition = mvPosition
          mvToken1 = ""
          mvToken2 = ""
          mvPosition = 1
          Do
            If mvEnv.Connection.IsUnicodeField(pCC.SubAttribute) Then
              vToken2 = NextToken(pSubsidiaryValue, TokenActionTypes.tatSubstitute, True)
            Else
              vToken2 = NextToken(pSubsidiaryValue, TokenActionTypes.tatSubstitute)
            End If
            If vToken2 <> "" Then
              If mvToken2 <> "" Then
                If vMainToken2 = "" Then
                  'vSQL = "SELECT FROM " & vTable & " WHERE " & vMainAttr & " = " & vMainToken1 & " AND " & vSubsidiaryAttr & " BETWEEN " & mvToken1 & " AND " & mvToken2
                  vSQL = "SELECT * FROM " & vTable & " WHERE " & vMainAttr & " = '" & vMainToken1 & "' AND " & vSubsidiaryAttr & " BETWEEN '" & mvToken1 & "' AND '" & mvToken2 & "'"
                Else
                  'vSQL = "SELECT FROM " & vTable & " WHERE " & vMainAttr & " BETWEEN " & vMainToken1 & " AND " & vMainToken2 & " AND " & vSubsidiaryAttr & " BETWEEN " & mvToken1 & " AND " & mvToken2 & " ORDER BY " & vSubsidiaryAttr
                  vSQL = "SELECT * FROM " & vTable & " WHERE " & vMainAttr & " BETWEEN '" & vMainToken1 & "' AND '" & vMainToken2 & "' AND '" & vSubsidiaryAttr & "' BETWEEN '" & mvToken1 & "' AND '" & mvToken2 & "' ORDER BY '" & vSubsidiaryAttr & "'"
                End If
                vRecordSet = mvConn.GetRecordSet(vSQL)
                While vRecordSet.Fetch() = True
                  If vSubAttrType <> "C" Then
                    If vList = "'" Then
                      vList = vRecordSet.Fields(vSubsidiaryAttr).Value
                    Else
                      vList = vList & "," & vRecordSet.Fields(vSubsidiaryAttr).Value
                    End If
                  Else
                    If vList = "'" Then
                      vList = vList & vRecordSet.Fields(vSubsidiaryAttr).Value & "'"
                    Else
                      vList = vList & ",'" & vRecordSet.Fields(vSubsidiaryAttr).Value & "'"
                    End If
                  End If
                End While
                vRecordSet.CloseRecordSet()
              End If
            End If
          Loop While vToken2 <> ""
          mvPosition = vMainPosition
          mvToken1 = ""
          mvToken2 = ""
        End If
      Loop While vToken <> ""

      If vList <> "'" Then
        mvSingleWhere = mvSingleWhere & vList
        BuildSubsidiaryList = True
      End If
    End Function
    Public Sub CopyDetails(ByRef pOriginalSetNumber As Integer, ByRef pOriginalRevision As Integer, ByRef pNewSetNumber As Integer, ByRef pNewRevision As Integer, ByVal pUnique As Boolean, Optional ByVal pTempToHold As Boolean = True, Optional ByVal pTempSourceTable As String = "", Optional ByVal pTempTargetTable As String = "")
      Dim vSQL As String
      Dim vRecords As Integer
      Dim vAttrs As String = ""
      Dim vSourceTable As String
      Dim vTargetTable As String
      Dim vSourceAttrs As String
      Dim vTargetAttrs As String
      Dim vNotIn As String

      If pTempToHold Then
        vSourceTable = mvSelectionTable
        vTargetTable = mvSelectionSetTable
      Else
        If pTempSourceTable.Length > 0 Then
          vSourceTable = pTempSourceTable
        Else
          vSourceTable = mvSelectionSetTable
        End If
        If pTempTargetTable.Length > 0 Then
          vTargetTable = pTempTargetTable
        Else
          vTargetTable = mvSelectionTable
        End If
      End If

      If mvMasterAttribute <> mvContactAttribute Then vAttrs = mvMasterAttribute & ", "
      vAttrs = vAttrs & mvContactAttribute & ", " & mvAddressAttribute
      If Left(vSourceTable, 9) = "selected_" Then
        'the source table is a selected_* table and doesn't have the address_number_2 attribute
        'so need to use the address_number attribute twice
        vSourceAttrs = vAttrs & ", " & mvAddressAttribute
        vTargetAttrs = vAttrs & ", address_number_2"
      ElseIf Left(vTargetTable, 9) = "selected_" Then
        'the target table is a selected_* table and doesn't have the address_number_2 attribute
        'so don't need to worry about the address_number_2 attribute
        vSourceAttrs = vAttrs
        vTargetAttrs = vAttrs
      Else
        'not dealing w/ a selected_* table so both tables will have the address_number_2 attribute
        vSourceAttrs = vAttrs & ", address_number_2"
        vTargetAttrs = vAttrs & ", address_number_2"
      End If

      vSQL = "INSERT INTO " & vTargetTable & " (selection_set, revision, " & vTargetAttrs & ")"
      vSQL = vSQL & " SELECT DISTINCT " & pNewSetNumber & ", " & pNewRevision & ", " & vSourceAttrs ' BR18523 - SELECT DISTINCT as as unique index will be created on the data.
      vSQL = vSQL & " FROM " & vSourceTable & " WHERE selection_set = " & pOriginalSetNumber
      If pOriginalRevision > 0 Then
        vSQL = vSQL & " AND revision = " & pOriginalRevision
      End If
      If pUnique Then
        vNotIn = " AND %1 NOT IN (SELECT %1 FROM " & vTargetTable & " WHERE selection_set = " & pNewSetNumber & " AND revision = " & pNewRevision & " )"
        'If pUnique = MsgBoxResult.No Then
        '  vSQL = vSQL & Replace(vNotIn, "%1", mvContactAttribute)
        'Else
        vSQL = vSQL & Replace(vNotIn, "%1", mvAddressAttribute)
        vSQL = vSQL & Replace(vNotIn, "%1", mvContactAttribute)
        'End If
      End If
      mvConn.LogMailSQL(vSQL)
      vRecords = mvConn.ExecuteSQL(vSQL)
    End Sub
    Private Sub CreateWorkRecord(ByRef pCC As CriteriaContext, ByRef pItem As CriteriaItems, ByRef pVariable As String)
      'Determines if a variable exists
      'Determines if variable has already been processed
      'If so ensures that it is not used in multiple ways.
      Dim vIndex As Integer
      Dim vFound As Boolean
      Dim vAttr As String = ""
      Dim vVarCriteria As VariableCriteria = Nothing
      Dim vValue As String = ""
      Dim vVariables() As String
      Dim vValues() As String
      Dim vSeparator As String

      If mvVariableParameters.Length > 0 Then
        vVariables = Split(mvVariableParameters, "|")
        For vIndex = 0 To UBound(vVariables)
          If Len(vVariables(vIndex)) > 0 Then
            'Determine which character separates the variable name from the variable value.
            'The old format was: variable name+variable value.  The new format is: variable name=variable value.
            'If the variable name is not $TODAY and "=" doesn't appears in either the name or the value, then assume the old format.
            'Other assume the new format.
            If UCase(Mid(vVariables(vIndex), 6)) <> "$TODAY" And InStr(vVariables(vIndex), "=") = 0 Then
              vSeparator = "+"
            Else
              vSeparator = "="
            End If
            vValues = Split(vVariables(vIndex), vSeparator)
            If vValues(0) = pVariable Then 'index 0 will be the variable name
              vValue = vValues(1) 'index 1 will be the variable value
              Exit For
            End If
          End If
        Next
      End If

      Select Case pItem
        Case CriteriaItems.citMainValue
          vAttr = "main_value"
        Case CriteriaItems.citSubValue
          vAttr = "subsidiary_value"
        Case CriteriaItems.citPeriod
          vAttr = "period"
      End Select

      'Determine if variable already exists in list
      For Each vVarCriteria In mvVariableCriteria
        With vVarCriteria
          If .VariableName = pVariable Then
            If .Valid And Left(.VariableName, 6).ToUpper <> "$TODAY" Then
              If .SearchArea_Renamed <> pCC.SearchArea Then
                RaiseError(DataAccessErrors.daeVarInMultipleAreas, pVariable)
              Else
                If .AttributeName <> vAttr Then RaiseError(DataAccessErrors.daeVarUsedMultipleWays, pVariable, .SearchArea_Renamed)
              End If
            End If
            vFound = True
            Exit For
          End If
        End With
      Next vVarCriteria
      If Not vFound Then
        vVarCriteria = New VariableCriteria
        mvVariableCriteria.Add(vVarCriteria)
      End If
      With vVarCriteria
        .VariableName = pVariable
        If .AttributeName <> vAttr Then .Value = ""
        .AttributeName = vAttr
        If .SearchArea_Renamed <> pCC.SearchArea Then .Value = ""
        .SearchArea_Renamed = pCC.SearchArea
        If vValue.Length > 0 Then .Value = vValue
        .ValTable = pCC.ValidationTable
        vVarCriteria.TableName = pCC.TableName
        Select Case vAttr
          Case "main_value"
            .ColumnHeading = pCC.MainValueHeading
            .ValAttribute = pCC.MainValidationAttribute
            .DataType = pCC.MainDataType
            .MainAttr = pCC.MainAttribute
          Case "subsidiary_value"
            .ColumnHeading = pCC.SubValueHeading
            .ValAttribute = pCC.SubValidationAttribute
            .MainAttr = pCC.MainValidationAttribute
            .MainValue = pCC.MainValue
            .DataType = pCC.SubDataType
          Case Else
            .ColumnHeading = "Period"
            .MainAttr = pCC.FromAttribute
            .DataType = "D"
        End Select
        .Pattern = pCC.Pattern
        .Valid = True
      End With
    End Sub
    Private Sub DedupAddresses(ByVal pSelectionSet As Integer, ByVal pRevision As Integer)
      Dim vContactNumber As Integer
      Dim vAddressNumber As Integer
      Dim vPreviousContact As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSelectionSet)
      vWhereFields.Add("revision", CDBField.FieldTypes.cftLong, pRevision)
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong)
      vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong)

      vRecordSet = mvConn.GetRecordSet("SELECT contact_number,address_number FROM " & mvSelectionTable & " WHERE selection_set = " & pSelectionSet & " AND revision = " & pRevision & " AND contact_number IN (SELECT contact_number FROM " & mvSelectionTable & " WHERE selection_set = " & pSelectionSet & " AND revision = " & pRevision & " GROUP BY contact_number HAVING COUNT (*) > 1) ORDER BY contact_number,address_number")
      While vRecordSet.Fetch() = True
        vContactNumber = CInt(vRecordSet.Fields("contact_number").Value)
        vAddressNumber = CInt(vRecordSet.Fields("address_number").Value)
        If vContactNumber = vPreviousContact Then
          vWhereFields(3).Value = CStr(vContactNumber)
          vWhereFields(4).Value = CStr(vAddressNumber)
          mvConn.DeleteRecords(mvSelectionTable, vWhereFields, False)
        End If
        vPreviousContact = vContactNumber
      End While
      vRecordSet.CloseRecordSet()
    End Sub
    Public Overloads Sub DeleteSelection(ByRef pSetNumber As Integer, ByRef pRevision As Integer, Optional ByVal pDeleteTable As Boolean = False, Optional ByVal pMasterAttribute As String = "", Optional ByVal pMasterAttributeValue As Integer = 0)
      Dim vWhereFields As New CDBFields

      If mvConn.AttributeExists(mvSelectionTable, "selection_set") Then
        If pDeleteTable Then
          mvConn.ExecuteSQL("DROP TABLE " & mvSelectionTable)
        Else
          vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSetNumber)
          If pRevision > 0 Then vWhereFields.Add("revision", CDBField.FieldTypes.cftLong, pRevision)
          If pMasterAttribute.Length > 0 Then vWhereFields.Add(pMasterAttribute, CDBField.FieldTypes.cftInteger, pMasterAttributeValue)
          mvConn.DeleteRecords(mvSelectionTable, vWhereFields, False)
        End If
      End If
    End Sub

    Public Overloads Sub DeleteSelection(ByRef pSetNumber As Integer, ByRef pRevision As Integer, ByVal pDeleteTable As Boolean, ByVal pMasterAttribute As String, pMasterAttributeValuesCSV As String) ' BR17119
      Dim vWhereFields As New CDBFields

      If mvConn.AttributeExists(mvSelectionTable, "selection_set") Then
        If pDeleteTable Then
          mvConn.ExecuteSQL("DROP TABLE " & mvSelectionTable)
        Else
          vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSetNumber)
          If pRevision > 0 Then vWhereFields.Add("revision", CDBField.FieldTypes.cftLong, pRevision)
          If pMasterAttribute.Length > 0 Then
            vWhereFields.Add(pMasterAttribute, CDBField.FieldTypes.cftInteger, pMasterAttributeValuesCSV, CDBField.FieldWhereOperators.fwoInOrEqual)
          End If
          mvConn.DeleteRecords(mvSelectionTable, vWhereFields, False)
        End If
      End If
    End Sub
    Private Function DoMatch(ByRef pSrc As String, ByRef pList As String) As Boolean
      Dim vPosition As Integer
      Dim vEndpos As Integer
      Dim vSubstr As String
      Dim vResult As Boolean

      vPosition = 1
      vResult = False
      Do
        vEndpos = InStr(vPosition, pList, "|")
        If vEndpos > 0 Then
          vSubstr = Mid(pList, vPosition, vEndpos - vPosition)
          vPosition = vEndpos + 1
        Else
          vSubstr = Mid(pList, vPosition)
        End If
        If pSrc = vSubstr Then
          vResult = True
        End If
      Loop While (vEndpos > 0) And (vResult <> True)
      DoMatch = vResult
    End Function
    Private Function ExtractToken(ByRef pString As String, ByRef pLength As Integer) As String
      Dim vChar As String
      Dim vToken As String = ""

      If mvTokenDelimiter = "'" Or mvTokenDelimiter = """" Then mvPosition = mvPosition + 1
      vChar = Mid(pString, mvPosition, 1)
      While Not DoMatch(vChar, mvTokenDelimiter) And mvPosition < pLength
        vToken = vToken & vChar
        mvPosition = mvPosition + 1
        vChar = Mid(pString, mvPosition, 1)
      End While
      If mvPosition = pLength Then
        If mvTokenDelimiter = " |," Then
          vToken = vToken & vChar
        Else
          If mvTokenDelimiter <> vChar Then RaiseError(DataAccessErrors.daeNoDelimiter)
        End If
        mvPosition = mvPosition + 1
      End If
      If vToken = "" Then
        RaiseError(DataAccessErrors.daeNullToken, Mid(pString, mvPosition))
      End If
      If mvPosition <= pLength Then
        If vChar <> "," Then mvPosition = mvPosition + 1
      End If
      ExtractToken = vToken
    End Function

    Private Sub GetCriteriaDetails(ByRef pCriteriaSet As Integer, ByRef pOrderBySequence As Boolean)
      Dim vSQL As String
      Dim vRecordSet As CDBRecordSet
      Dim vCC As CriteriaContext
      Dim vID As Integer
      Dim vAttrs As String

      vAttrs = "criteria_set,csd.search_area,csd.c_o,i_e,main_value,subsidiary_value,main_data_type,subsidiary_data_type,period,counted,and_or,left_parenthesis,right_parenthesis"
      vAttrs = vAttrs & ",sc.validation_table,main_attribute,main_validation_attribute, subsidiary_attribute,"
      vAttrs = vAttrs & mvConn.DBAttrName("subsidiary_validation_attribute") & ",main_value_heading,subsidiary_value_heading"
      vAttrs = vAttrs & ",to_attribute,from_attribute,indexed,special,sc.table_name,address_link,geographic,csd.sequence_number,nulls_allowed,pattern"

      vSQL = "SELECT " & vAttrs & " FROM criteria_set_details csd, selection_control sc, maintenance_attributes ma WHERE csd.criteria_set = " & pCriteriaSet
      vSQL = vSQL & " AND csd.search_area = sc.search_area AND csd.c_o = sc.c_o AND sc.application_name = '" & mvApplication & "'"
      vSQL = vSQL & " AND sc.table_name = ma.table_name AND sc.main_attribute = ma.attribute_name"
      If pOrderBySequence Then
        vSQL = vSQL & " ORDER BY csd.sequence_number"
      Else
        vSQL = vSQL & " ORDER BY csd.i_e desc, csd.counted"
      End If
      mvCriteriaContexts = New Collection
      mvConCriteriaCount = 0
      mvOrgCriteriaCount = 0
      vRecordSet = mvConn.GetRecordSet(vSQL)
      'If vRecordSet.Fetch() <> rssOK Then RaiseError daeNoCriteria
      While vRecordSet.Fetch() = True
        vID = vID + 1
        vCC = New CriteriaContext
        vCC.InitFromRecordSet(mvEnv, mvConn, vRecordSet, vID)
        mvCriteriaContexts.Add(vCC, CStr(vID))
        If vCC.Contacts Then
          mvConCriteriaCount = mvConCriteriaCount + 1
        Else
          mvOrgCriteriaCount = mvOrgCriteriaCount + 1
        End If
      End While
      vRecordSet.CloseRecordSet()
    End Sub
    Private Function GetVariableValue(ByRef pVariable As String) As String
      Dim vVarCriteria As VariableCriteria

      For Each vVarCriteria In mvVariableCriteria
        If vVarCriteria.Valid Then
          If vVarCriteria.VariableName = pVariable Then
            Return vVarCriteria.Value
          End If
        End If
      Next vVarCriteria
      Return ""
    End Function
    Public Function GetMailingCode(ByRef pType As MailingTypes) As String

      Select Case pType
        Case MailingTypes.mtyCampaigns
          Return "CA"
        Case MailingTypes.mtyDirectDebits
          Return "DD"
        Case MailingTypes.mtyGeneralMailing
          Return "GM"
        Case MailingTypes.mtyStandardExclusions
          Return "GM"
        Case MailingTypes.mtyMembers
          Return "MM"
        Case MailingTypes.mtyMembershipCards
          Return "MC"
        Case MailingTypes.mtyPerformanceAnalyser
          Return "PA"
        Case MailingTypes.mtyPayers, MailingTypes.mtyRenewalsAndReminders
          Return "PM"
        Case MailingTypes.mtyScoringAnalyser
          Return "SA"
        Case MailingTypes.mtySubscriptions
          Return "SL"
        Case MailingTypes.mtyStandingOrders
          Return "SO"
        Case MailingTypes.mtySelectionTester
          Return "ST"
        Case MailingTypes.mtyStandingOrderCancellation
          Return "BC"
        Case MailingTypes.mtyEventAttendees
          Return "EA"
        Case MailingTypes.mtyEventBookings
          Return "EB"
        Case MailingTypes.mtyEventPersonnel
          Return "EP"
        Case MailingTypes.mtyMemberFulfilment
          Return "MF"
        Case MailingTypes.mtyNonMemberFulfilment
          Return "NF"
        Case MailingTypes.mtyGAYECancellation
          Return "PC"
        Case MailingTypes.mtyGAYEPledges
          Return "GP"
        Case MailingTypes.mtyNameGathering
          Return "NG"
        Case MailingTypes.mtySaleOrReturn
          Return "SR"
        Case MailingTypes.mtyEventSponsors
          Return "ES"
        Case MailingTypes.mtyIrishGiftAid
          Return "IG"
        Case MailingTypes.mtyExamBookings
          Return "XB"
        Case MailingTypes.mtyExamCandidates
          Return "XC"
        Case Else
          RaiseError(DataAccessErrors.daeUnknownMailingType)
          Return ""       'Fix compiler warning
      End Select
    End Function
    Public Function GetMailingDescription(ByRef pType As MailingTypes) As String

      Select Case pType
        Case MailingTypes.mtyCampaigns
          Return "Campaign Appeal Mailing"
        Case MailingTypes.mtyDirectDebits
          Return "Direct Debit Mailing"
        Case MailingTypes.mtyGeneralMailing
          Return "Selection Manager"
        Case MailingTypes.mtyStandardExclusions
          Return "Standard Exclusions"
        Case MailingTypes.mtyMembers
          Return "Member Mailing"
        Case MailingTypes.mtyMembershipCards
          Return "Membership Cards"
        Case MailingTypes.mtyPerformanceAnalyser
          Return "Performance Analyser"
        Case MailingTypes.mtyPayers, MailingTypes.mtyRenewalsAndReminders
          Return "Payer Mailing"
        Case MailingTypes.mtyScoringAnalyser
          Return "Scoring Analyser"
        Case MailingTypes.mtySubscriptions
          Return "Subscription Mailing"
        Case MailingTypes.mtyStandingOrders
          Return "Standing Order Mailing"
        Case MailingTypes.mtySelectionTester
          Return "Selection Tester"
        Case MailingTypes.mtyStandingOrderCancellation
          Return "Standing Order Cancellation"
        Case MailingTypes.mtyEventAttendees
          Return "Event Delegates Mailing"
        Case MailingTypes.mtyEventBookings
          Return "Event Bookings Mailing"
        Case MailingTypes.mtyEventPersonnel
          Return "Event Personnel Mailing"
        Case MailingTypes.mtyMemberFulfilment
          Return "New Member Fulfilment"
        Case MailingTypes.mtyNonMemberFulfilment
          Return "Non-Member Fulfilment"
        Case MailingTypes.mtyGAYECancellation
          Return "Payroll Giving Pledges Cancellation"
        Case MailingTypes.mtyGAYEPledges
          Return "Payroll Giving Pledge Mailing"
        Case MailingTypes.mtyNameGathering
          Return "Name Gathering Mailing"
        Case MailingTypes.mtySaleOrReturn
          Return "Sale Or Return Mailing"
        Case MailingTypes.mtyEventSponsors
          Return "Event Sponsors Mailing"
        Case MailingTypes.mtyIrishGiftAid
          Return "Irish Gift Aid Mailing"
        Case MailingTypes.mtyExamBookings
          Return "Exam Bookings Mailing"
        Case MailingTypes.mtyExamCandidates
          Return "Exam Candidates Mailing"
        Case Else
          RaiseError(DataAccessErrors.daeUnknownMailingType)
          Return ""       'Fix compiler warning
      End Select
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pTypeCode As String, Optional ByVal pSelectionSet As Integer = 0, Optional ByVal pRAndR As Boolean = False, Optional ByVal pFromAppealOrSegment As Boolean = False, Optional ByVal pDisplayOrgSelection As Boolean = True)
      Dim vRS As CDBRecordSet

      mvEnv = pEnv
      mvConn = pEnv.Connection
      mvVariableCriteria = New Collection
      mvLinkToContacts = True
      mvDisplayOrgSelection = pDisplayOrgSelection 'True
      mvMasterHasAddress = True
      Select Case pTypeCode
        Case "CA"
          mvMailingType = MailingTypes.mtyCampaigns
          mvMasterTable = "contact_header"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Campaign Appeal Mailing"
        Case "DD"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyDirectDebits
          mvMasterTable = "direct_debits"
          mvMasterAttribute = "direct_debit_number"
          mvSelectionSetTable = "selected_direct_debits"
          mvCaption = "Direct Debit Mailing"
        Case "GM"
          mvMailingType = MailingTypes.mtyGeneralMailing
          mvMasterTable = "contacts"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Selection Manager"
        Case "MC", "MM"
          mvLinkToContacts = False
          If pTypeCode = "MM" Then
            mvCaption = "Member Mailing"
            mvMailingType = MailingTypes.mtyMembers
          Else
            mvCaption = "Membership Cards"
            pTypeCode = "MM"
            mvMailingType = MailingTypes.mtyMembershipCards
            Select Case pEnv.GetConfig("me_card_production_select")
              Case "PAYMENT_REQUIRED"
                mvCardProductionType = MembershipCardProductionTypes.mcpPaymentRequired
              Case "PAID_OR_AUTO"
                mvCardProductionType = MembershipCardProductionTypes.mcpAutoOrPaid
              Case Else
                mvCardProductionType = MembershipCardProductionTypes.mcpDefault
            End Select
          End If
          mvMasterTable = "members"
          mvMasterAttribute = "membership_number"
          mvSelectionSetTable = "selected_members"
        Case "NG"
          mvMailingType = MailingTypes.mtyNameGathering
          mvMasterTable = "contact_incentive_responses"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Name Gathering Mailing"
        Case "PA"
          mvMailingType = MailingTypes.mtyPerformanceAnalyser
          mvMasterTable = "contacts"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Performance Analyser"
        Case "PM"
          mvLinkToContacts = False
          If pRAndR Then
            mvMailingType = MailingTypes.mtyRenewalsAndReminders
            mvCaption = "Renewals And Reminders"
          Else
            mvMailingType = MailingTypes.mtyPayers
            mvCaption = "Payer Mailing"
          End If
          mvMasterTable = "orders"
          mvMasterAttribute = "order_number"
          mvSelectionSetTable = "selected_orders"
        Case "SA"
          mvMailingType = MailingTypes.mtyScoringAnalyser
          mvMasterTable = "contacts"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Scoring Analyser"
        Case "SE"
          mvMailingType = MailingTypes.mtyStandardExclusions
          mvMasterTable = "contacts"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Standard Exclusions"
        Case "SL"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtySubscriptions
          mvMasterTable = "subscriptions"
          mvMasterAttribute = "subscription_number"
          mvSelectionSetTable = "selected_subscriptions"
          mvCaption = "Subscription Mailing"
        Case "SO"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyStandingOrders
          mvMasterTable = "bankers_orders"
          mvMasterAttribute = "bankers_order_number"
          mvSelectionSetTable = "selected_bankers_orders"
          mvCaption = "Standing Order Mailing"
        Case "ST"
          mvMailingType = MailingTypes.mtySelectionTester
          mvMasterTable = "contact_header"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Selection Tester"
        Case "BC"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyStandingOrderCancellation
          mvMasterTable = "bankers_orders"
          mvMasterAttribute = "bankers_order_number"
          mvSelectionSetTable = "selected_bankers_orders"
          mvCaption = "Standing Order Cancellation"
        Case "EA"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyEventAttendees
          mvMasterTable = "delegates"
          mvMasterAttribute = "event_delegate_number"
          mvSelectionSetTable = "selected_event_attendees"
          mvCaption = "Event Delegates Mailing"
        Case "EB"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyEventBookings
          mvMasterTable = "event_bookings"
          mvMasterAttribute = "booking_number"
          mvSelectionSetTable = "selected_event_bookings"
          mvCaption = "Event Bookings Mailing"
        Case "EP"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyEventPersonnel
          mvMasterTable = "event_personnel"
          mvMasterAttribute = "event_personnel_number"
          mvSelectionSetTable = "selected_event_personnel"
          mvCaption = "Event Personnel Mailing"
        Case "MF"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyMemberFulfilment
          mvMasterTable = "new_orders"
          mvMasterAttribute = "order_number"
          mvSelectionSetTable = "selected_orders"
          mvCaption = "New Member Fulfilment"
          mvDisplayOrgSelection = False
        Case "NF"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyNonMemberFulfilment
          mvMasterTable = "new_orders"
          mvMasterAttribute = "order_number"
          mvSelectionSetTable = "selected_orders"
          mvCaption = "Non-Member Fulfilment"
          mvDisplayOrgSelection = False
        Case "SR"
          mvMailingType = MailingTypes.mtySaleOrReturn
          mvMasterTable = "contacts"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts"
          mvCaption = "Sale Or Return Mailing"
        Case "PC", "GP"
          mvMasterTable = "gaye_pledges"
          mvMasterAttribute = "gaye_pledge_number"
          mvSelectionSetTable = "selected_gaye_pledges"
          If pTypeCode = "PC" Then
            mvCaption = "Payroll Giving Pledges Cancellation"
            mvMailingType = MailingTypes.mtyGAYECancellation
          Else
            mvCaption = "Payroll Giving Pledge Mailing"
            mvMailingType = MailingTypes.mtyGAYEPledges
            mvLinkToContacts = False
          End If
          mvMasterHasAddress = False
        Case "ES"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyEventSponsors
          mvMasterTable = "sundry_costs"
          mvMasterAttribute = "sundry_cost_number"
          mvSelectionSetTable = "selected_event_sponsors"
          mvCaption = "Event Sponsors Mailing"
        Case "IG"
          mvMailingType = MailingTypes.mtyIrishGiftAid
          mvMasterTable = "contact_performances"
          mvMasterAttribute = "contact_number"
          mvSelectionSetTable = "selected_contacts" 'certificates"
          mvCaption = "Irish Gift Aid Mailing"
        Case "XB"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyExamBookings
          mvMasterTable = "exam_bookings"
          mvMasterAttribute = "exam_booking_id"
          mvSelectionSetTable = "selected_exam_bookings"
          mvCaption = "Exam Bookings Mailing"
        Case "XC"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyExamCandidates
          mvMasterTable = "exam_booking_units"
          mvMasterAttribute = "exam_booking_unit_id"
          mvSelectionSetTable = "selected_exam_candidates"
          mvCaption = "Exam Candidates Mailing"
        Case "XR"
          mvLinkToContacts = False
          mvMailingType = MailingTypes.mtyExamCertificates
          mvMasterTable = "vcontactstudentunitheader"
          mvMasterAttribute = "exam_student_unit_header_id"
          mvSelectionSetTable = "selected_student_unit_headers"
          mvCaption = "Exam Certificates"
        Case Else
          RaiseError(DataAccessErrors.daeUnknownMailingType)
      End Select
      'Set Defaults for various class variables
      mvApplication = pTypeCode
      mvContactAttribute = "contact_number"
      mvOrgTable = "organisations"
      mvOrgAttribute = "organisation_number"
      mvPositionTable = "contact_positions"
      mvContactAddressTable = "contact_addresses"
      mvOrgAddressTable = "organisation_addresses"
      mvAddressAttribute = "address_number"
      mvIndexNoOptimise = False
      mvOrgMailTo = OrgSelectContact.oscAllEmployees
      mvOrgMailWhere = OrgSelectAddress.osaOrganisationAddress
      mvSelectionTable = "smcam_smapp"
      If pSelectionSet > 0 Then mvSelectionTable = mvSelectionTable & "_" & pSelectionSet
      If pFromAppealOrSegment Then mvAppealMailing = True

      mvMasterAttrTables = New CDBCollection
      mvOrgAttrTables = New CDBCollection
      vRS = mvConn.GetRecordSet("SELECT DISTINCT table_name, attribute_name FROM maintenance_attributes WHERE table_name NOT LIKE 'ext_%' AND (attribute_name LIKE '%" & mvMasterAttribute & "%' OR attribute_name = '" & mvOrgAttribute & "')")
      With vRS
        While .Fetch() = True
          If InStr(.Fields.Item(2).Value, mvMasterAttribute) > 0 Then
            If Not mvMasterAttrTables.Exists((.Fields.Item(1).Value)) Then mvMasterAttrTables.Add((.Fields.Item(1).Value), (.Fields.Item(1).Value))
          Else
            If Not mvOrgAttrTables.Exists((.Fields.Item(1).Value)) Then mvOrgAttrTables.Add((.Fields.Item(1).Value), (.Fields.Item(1).Value))
          End If
        End While
        .CloseRecordSet()
      End With

      mvTableAliases = New CDBCollection
    End Sub
    Public Sub InitVariableCriteria(ByRef pCriteriaSet As Integer)
      Dim vCC As CriteriaContext

      GetCriteriaDetails(pCriteriaSet, True)
      For Each vCC In mvCriteriaContexts
        LocateVariables(vCC, CriteriaItems.citMainValue)
        LocateVariables(vCC, CriteriaItems.citSubValue)
        LocateVariables(vCC, CriteriaItems.citPeriod)
      Next vCC
    End Sub
    Private Function LinkToAddress() As String
      Dim vSelectTable As String

      If mvOGeographic = "" Then
        If mvCGeographic <> "" Then
          LinkToAddress = mvCGeographic
        Else
          If mvLastOTable = "" Then
            If mvCAddressLink <> "" Then
              LinkToAddress = mvCAddressLink
            Else
              If mvContactTableAlias <> "" Then
                LinkToAddress = mvContactTableAlias
              Else
                ' join to master table
                mvTableNumber = mvTableNumber + 1
                vSelectTable = "table" & CStr(mvTableNumber)
                mvTableList = mvTableList & ", " & mvMasterTable & " " & vSelectTable
                If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvMasterTable, vSelectTable)
                mvWhere = mvWhere & " AND " & vSelectTable & "." & mvMasterAttribute & " = " & mvLastCTable & "." & mvMasterAttribute
                If Len(mvMasterTableAlias) = 0 Then mvMasterTableAlias = vSelectTable
                LinkToAddress = vSelectTable
              End If
            End If
          Else
            mvTableNumber = mvTableNumber + 1
            vSelectTable = "table" & CStr(mvTableNumber)
            mvWhere = mvWhere & " AND " & mvLastOTable & "." & mvOrgAttribute & " = " & vSelectTable & ".organisation_number AND " & vSelectTable & ".historical = 'N' AND " & vSelectTable & "." & mvAddressAttribute & " = "
            mvTableList = mvTableList & ", " & mvOrgAddressTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvOrgAddressTable, vSelectTable)
            mvTableNumber = mvTableNumber + 1
            vSelectTable = "table" & CStr(mvTableNumber)
            mvWhere = mvWhere & vSelectTable & "." & mvAddressAttribute & " AND " & vSelectTable & ".historical = 'N' AND " & vSelectTable & "." & mvContactAttribute & " = " & mvLastCTable & "." & mvContactAttribute
            mvTableList = mvTableList & ", " & mvContactAddressTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvContactAddressTable, vSelectTable)
            LinkToAddress = vSelectTable
          End If
        End If
      Else
        mvTableNumber = mvTableNumber + 1
        vSelectTable = "table" & CStr(mvTableNumber)
        mvWhere = mvWhere & " AND " & mvOGeographic & "." & mvAddressAttribute & " = " & vSelectTable & "." & mvAddressAttribute & " AND " & vSelectTable & ".historical = 'N' AND " & vSelectTable & "." & mvContactAttribute & " = " & mvLastCTable & "." & mvContactAttribute
        mvTableList = mvTableList & ", " & mvContactAddressTable & " " & vSelectTable
        If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvContactAddressTable, vSelectTable)
        LinkToAddress = mvOGeographic
      End If
    End Function
    Private Sub LinkToContacts()
      Dim vSelectTable As String
      Dim vPos As Integer
      Dim vEnd As Integer

      'if only C type search areas ensure we have joined to the contacts table
      'and check the contact_type <> 'O'  ie force contact only
      If mvLinkToContacts Then
        If mvLastOTable = "" Then
          'C type search areas only
          vPos = InStr(mvTableList, " contacts table")
          If vPos <= 0 Then
            'add contacts table to the select and join to it from the last c table
            mvTableNumber = mvTableNumber + 1
            vSelectTable = "table" & CStr(mvTableNumber)
            mvTableList = mvTableList & ", contacts " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add("contacts", vSelectTable)
            mvWhere = mvWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = " & vSelectTable & "." & mvContactAttribute
            mvContactTableAlias = vSelectTable
          Else
            'we have already joined to contacts - find the contacts table alias
            vEnd = Len(mvTableList)
            vPos = vPos + 10
            'we are at the start of the alias name now
            mvContactTableAlias = Mid(mvTableList, vPos, 1)
            vPos = vPos + 1
            While Mid(mvTableList, vPos, 1) <> "," And Mid(mvTableList, vPos, 1) <> " " And vPos <= vEnd
              mvContactTableAlias = mvContactTableAlias & Mid(mvTableList, vPos, 1)
              vPos = vPos + 1
            End While
          End If
          mvWhere = mvWhere & " AND " & mvContactTableAlias & ".contact_type <> 'O'"
        End If
      End If
    End Sub
    Private Sub LinkToMaster()
      Dim vSelectTable As String
      Dim vPos As Integer
      Dim vEnd As Integer

      'ensure we have joined to the master table
      vPos = InStr(mvTableList, " " & mvMasterTable & " table")
      If vPos <= 0 Then
        mvTableNumber = mvTableNumber + 1
        vSelectTable = "table" & CStr(mvTableNumber)
        mvTableList = mvTableList & ", " & mvMasterTable & " " & vSelectTable
        If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvMasterTable, vSelectTable)
        mvWhere = mvWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = " & vSelectTable & "." & mvContactAttribute
        mvLastCTable = vSelectTable
        If mvContactTableAlias = "" Then mvContactTableAlias = vSelectTable
        If Len(mvMasterTableAlias) = 0 Then mvMasterTableAlias = vSelectTable
      Else
        If mvContactTableAlias = "" Then
          'bug in the code before this bit as the master table is in the
          'table list but the contact_table_alias has not been set.
          'Yes he knows the attribute is named incorrectly as it should be
          'gv_zse_master_table_alias but he can't fix this at the moment.
          vEnd = Len(mvTableList)
          vPos = vPos + Len(mvMasterTable) + 2
          'we are at the start of the alias name now
          mvContactTableAlias = Mid(mvTableList, vPos, 1)
          vPos = vPos + 1
          While Mid(mvTableList, vPos, 1) <> "," And Mid(mvTableList, vPos, 1) <> " " And vPos <= vEnd
            mvContactTableAlias = mvContactTableAlias & Mid(mvTableList, vPos, 1)
            vPos = vPos + 1
          End While
        End If
      End If
    End Sub
    Private Sub LinkToPositions()
      Dim vSelectTable As String
      Dim vPositionTableAlias As String
      Dim vOrgAttr As String

      If mvLastCTable = "" Then
        'only O type search areas have been used
        'Joint the the contact positions table to get the appropriate contacts

        mvTableNumber = mvTableNumber + 1
        vSelectTable = "table" & CStr(mvTableNumber)
        mvTableList = mvTableList & ", " & mvPositionTable & " " & vSelectTable
        If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvPositionTable, vSelectTable)
        vPositionTableAlias = vSelectTable

        vOrgAttr = mvOrgAttribute
        If mvOrgMailTo > 0 Then
          If Not mvOrgAttrTables.Exists(CStr(mvTableAliases(mvLastOTable))) Then vOrgAttr = mvContactAttribute
        End If

        Select Case mvOrgMailTo
          Case OrgSelectContact.oscAllEmployees 'Mail to all mailable employees
            mvWhere = mvWhere & " AND " & mvLastOTable & "." & vOrgAttr & " = " & vSelectTable & ".organisation_number AND " & vSelectTable & ".organisation_number <> " & vSelectTable & "." & mvContactAttribute & " AND " & mvConn.DBSpecialCol(vSelectTable, "current") & " = 'Y' AND " & vSelectTable & "." & "mail = 'Y'"
          Case OrgSelectContact.oscDefaultContact 'Just join to position table, restrict to dfault below
            mvWhere = mvWhere & " AND " & mvLastOTable & "." & vOrgAttr & " = " & vSelectTable & ".organisation_number"
          Case Else 'Mail only the dummy contact
            mvWhere = mvWhere & " AND " & mvLastOTable & "." & vOrgAttr & " = " & vSelectTable & ".organisation_number AND " & vSelectTable & ".organisation_number = " & vSelectTable & "." & mvContactAttribute
        End Select

        mvLastCTable = vSelectTable
        mvLastOTable = vSelectTable
        If mvCAddressLink = "" Then mvCAddressLink = vSelectTable

        'ensure we have joined to the master table
        If InStr(mvTableList, mvMasterTable) <= 0 Then
          mvTableNumber = mvTableNumber + 1
          vSelectTable = "table" & CStr(mvTableNumber)
          mvTableList = mvTableList & ", " & mvMasterTable & " " & vSelectTable
          If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvMasterTable, vSelectTable)
          mvWhere = mvWhere & " AND " & mvLastCTable & "." & mvContactAttribute & " = " & vSelectTable & "." & mvContactAttribute
          mvLastCTable = vSelectTable
          If mvContactTableAlias = "" Then mvContactTableAlias = vSelectTable
          If Len(mvMasterTableAlias) = 0 Then mvMasterTableAlias = vSelectTable
        End If

        'force mailing to default contact if requested */
        If mvOrgMailTo = OrgSelectContact.oscDefaultContact Then
          'we need the organisation table in the select so add it if ness
          If mvOrgTableAlias = "" Then
            mvTableNumber = mvTableNumber + 1
            vSelectTable = "table" & CStr(mvTableNumber)
            mvTableList = mvTableList & ", " & mvOrgTable & " " & vSelectTable
            If Not mvTableAliases.Exists(vSelectTable) Then mvTableAliases.Add(mvOrgTable, vSelectTable)
            mvOrgTableAlias = vSelectTable
            mvWhere = mvWhere & " AND " & mvLastOTable & "." & mvOrgAttribute & " = " & vSelectTable & "." & mvOrgAttribute
          End If
          mvWhere = mvWhere & " AND " & vPositionTableAlias & "." & mvContactAttribute & " = " & mvOrgTableAlias & "." & mvContactAttribute
          If mvMasterAttribute <> mvContactAttribute Then mvContactTableAlias = vPositionTableAlias
        End If
      End If
    End Sub
    Private Sub LocateVariables(ByRef pCC As CriteriaContext, ByRef pItem As CriteriaItems)
      'Used to process any attribute of a criteria set details record trying to locate variables.
      Dim vValue As String = ""
      Dim vToken As String
      Dim vUnicode As Boolean = False

      mvPosition = 1
      Select Case pItem
        Case CriteriaItems.citMainValue
          vValue = pCC.MainValue
          vUnicode = mvEnv.Connection.IsUnicodeField(pCC.MainAttribute)
        Case CriteriaItems.citSubValue
          vValue = pCC.SubValue
          vUnicode = mvEnv.Connection.IsUnicodeField(pCC.SubAttribute)
        Case CriteriaItems.citPeriod
          vValue = pCC.Period
      End Select
      Do

        vToken = NextToken(vValue, TokenActionTypes.tatLocate, vUnicode)
        If vToken <> "" Then
          If Left(mvToken1, 1) = "$" Then CreateWorkRecord(pCC, pItem, mvToken1)
          If Left(mvToken2, 1) = "$" Then CreateWorkRecord(pCC, pItem, mvToken2)
        End If
      Loop While vToken <> ""
    End Sub
    Private Function NextToken(ByRef pValue As String, ByRef pMode As TokenActionTypes, Optional ByVal pUnicode As Boolean = False) As String
      Dim vString As String
      Dim vChar As String
      Dim vWord As String
      Dim vLength As Integer
      Dim vPrevPosition As Integer

      If pValue = "" Or mvPosition > Len(pValue) Then
        NextToken = ""
      Else
        vLength = Len(pValue)
        mvToken1 = ""
        mvToken2 = ""
        vString = pValue
        vChar = Mid(vString, mvPosition, 1)
        If pUnicode = False And (LCase(vChar) < "a" Or LCase(vChar) > "z") And (vChar < "0" Or vChar > "9") And vChar <> "'" And vChar <> """" And vChar <> "$" And vChar <> "*" And vChar <> "_" Then
          If mvPosition = 1 Then
            RaiseError(DataAccessErrors.daeInvalidCharAtStart, Mid(vString, mvPosition, 1))
          Else
            RaiseError(DataAccessErrors.daeInvalidCharAfter, Mid(pValue, mvPosition - 1, 1))
          End If
        Else
          If vChar = "'" Or vChar = """" Then
            mvTokenDelimiter = vChar
          Else
            mvTokenDelimiter = " |,"
          End If
        End If
        vPrevPosition = mvPosition
        mvToken1 = ExtractToken(vString, vLength)
        'msd 24jun96
        If Left(mvToken1, 1) = "$" Then
          If pMode = TokenActionTypes.tatParseOnly Then
            If mvPosition < vLength Then
              vChar = Mid(vString, mvPosition, 1)
              If vChar = "t" Then
                RaiseError(DataAccessErrors.daeVariableInRange, mvToken1)
              Else
                SkipToToken(vString, vLength)
              End If
              mvToken1 = ExtractToken(vString, vLength)
            Else
              mvToken1 = ""
            End If
          Else
            If pMode = TokenActionTypes.tatSubstitute Then
              vString = Replace(vString, mvToken1, GetVariableValue(mvToken1))
              'msd 02sep96
              vChar = Left(vString, 1)
              If (LCase(vChar) < "a" Or LCase(vChar) > "z") And (vChar < "0" Or vChar > "9") And vChar <> "'" And vChar <> """" And vChar <> "*" Then
                If mvPosition = 1 Then
                  RaiseError(DataAccessErrors.daeInvalidCharAtStart, Mid(vString, mvPosition, 1))
                Else
                  RaiseError(DataAccessErrors.daeInvalidCharAfter, Mid(pValue, mvPosition - 1, 1))
                End If
              Else
                If vChar = "'" Or vChar = """" Then
                  mvTokenDelimiter = vChar
                Else
                  mvTokenDelimiter = " |,"
                End If
              End If
              pValue = vString
              mvPosition = vPrevPosition
              vLength = Len(vString)
              mvToken1 = ExtractToken(vString, vLength)
            End If
          End If
        End If
        If mvPosition > vLength Then
          NextToken = mvToken1
        Else
          vChar = Mid(vString, mvPosition, 1)
          If vChar = "," Then 'assume list or free text, so skip and return
            SkipToToken(vString, vLength)
          Else
            ' could be 'to' and therefore a range or spaces before a comma
            While vChar = " " And mvPosition < vLength
              mvPosition = mvPosition + 1
              vChar = Mid(vString, mvPosition, 1)
            End While
            If vChar = "t" Then
              vWord = "t"
              mvPosition = mvPosition + 1
              vChar = Mid(vString, mvPosition, 1)
              While vChar <> " " And vChar <> "'" And vChar <> """" And mvPosition < vLength
                vWord = vWord & vChar
                mvPosition = mvPosition + 1
                vChar = Mid(vString, mvPosition, 1)
              End While
              If vWord <> "to" Then
                RaiseError(DataAccessErrors.daeInvalidTo, Mid(vString, mvPosition - 1))
              Else
                While vChar = " " And mvPosition < vLength
                  mvPosition = mvPosition + 1
                  vChar = Mid(vString, mvPosition, 1)
                End While
                If mvPosition > vLength Then
                  RaiseError(DataAccessErrors.daeEndOfCriteria)
                Else
                  If (LCase(vChar) < "a" Or LCase(vChar) > "z") And (vChar < "0" Or vChar > "9") And vChar <> "'" And vChar <> """" And vChar <> "$" And vChar <> "*" Then
                    RaiseError(DataAccessErrors.daeInvalidCharAfter, Mid(pValue, mvPosition - 1, 1))
                  Else
                    If vChar = "'" Or vChar = """" Then
                      mvTokenDelimiter = vChar
                    Else
                      mvTokenDelimiter = " |,"
                    End If
                    ' found to so can extract second token
                    mvToken2 = ExtractToken(vString, vLength)
                    If Left(mvToken2, 1) = "$" Then
                      RaiseError(DataAccessErrors.daeVariableInRange, mvToken2)
                    Else
                      If mvPosition < vLength Then
                        vChar = Mid(vString, mvPosition, 1)
                        If vChar <> "," Then
                          RaiseError(DataAccessErrors.daeMissingDelimiterAfter, Mid(vString, mvPosition - 1))
                        Else
                          SkipToToken(vString, vLength)
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Else
              If vChar = "," Then ' assume list or free text, so skip and return
                SkipToToken(vString, vLength)
              Else
                RaiseError(DataAccessErrors.daeMissingDelimiterAfter, Left(vString, mvPosition - 1))
              End If
            End If
          End If
          NextToken = mvToken1
        End If
      End If
    End Function
    Private Sub ProcessDates(ByRef pCC As CriteriaContext, ByRef pTable As String)
      Dim vCount As Integer
      Dim vToken As String

      mvPosition = 1
      mvToken1 = ""
      mvToken2 = ""
      vCount = 0
      mvSingleWhere = mvSingleWhere & " ( "
      Do
        vToken = NextToken((pCC.Period), TokenActionTypes.tatSubstitute)
        If vToken <> "" Then
          vCount = vCount + 1
          If vCount > 1 Then mvSingleWhere = mvSingleWhere & " OR "
          If mvToken2 = "" Then
            If pCC.ToAttribute = "" Then
              mvSingleWhere = mvSingleWhere & pTable & "." & pCC.FromAttribute & mvConn.SQLLiteral("=", CDBField.FieldTypes.cftDate, mvToken1)
            Else
              If pCC.SearchArea = "groupdates" And pCC.TableName = "membership_groups" Then
                mvSingleWhere = mvSingleWhere & pTable & "." & pCC.FromAttribute & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, mvToken1)
              Else
                mvSingleWhere = mvSingleWhere & pTable & "." & pCC.FromAttribute & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, mvToken1) & " AND "
                If pCC.SearchArea = "role" Then mvSingleWhere = mvSingleWhere & " (( "
                mvSingleWhere = mvSingleWhere & pTable & "." & pCC.ToAttribute & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, mvToken1)

                If pCC.SearchArea = "role" Then
                  'The contact_roles.valid_to can be null so need to include those where valid_to is null
                  mvSingleWhere = mvSingleWhere & " ) OR ( "
                  mvSingleWhere = mvSingleWhere & pTable & "." & pCC.ToAttribute & " IS NULL )) "
                End If
              End If
            End If
          Else
            If pCC.ToAttribute = "" Then
              mvSingleWhere = mvSingleWhere & pTable & "." & pCC.FromAttribute & mvConn.SQLLiteral("BETWEEN", CDBField.FieldTypes.cftDate, mvToken1) & mvConn.SQLLiteral("AND", CDBField.FieldTypes.cftDate, mvToken2)
            Else
              mvSingleWhere = mvSingleWhere & pTable & "." & pCC.FromAttribute & mvConn.SQLLiteral(">=", CDBField.FieldTypes.cftDate, mvToken1) & " AND " & pTable & "." & pCC.ToAttribute & mvConn.SQLLiteral("<=", CDBField.FieldTypes.cftDate, mvToken2)
            End If
          End If
        End If
      Loop While vToken <> ""
      mvSingleWhere = mvSingleWhere & " ) "
    End Sub
    Private Sub ProcessMainValue(ByRef pCC As CriteriaContext, ByRef pValue As String, ByRef pSelectTable As String)
      Dim vListTokens As Integer
      Dim vListSet As String
      Dim vRangeTokens As Integer
      Dim vToken As String
      Dim vDelimiter As String
      Dim vMainAttr As String
      Dim vListArray() As String
      Dim vListItem As String
      Dim vIndex As Integer

      vDelimiter = If(pCC.MainDataType = "C", "'", "") 'Single quotes not valid for numerics in ACCESS or ODBC?
      mvPosition = 1
      mvToken1 = ""
      mvToken2 = ""
      vListSet = "'"
      vListTokens = 0
      vRangeTokens = 0
      vMainAttr = pCC.MainAttribute
      If mvConn.IsSpecialColumn(vMainAttr) Then vMainAttr = mvConn.DBSpecialCol("", vMainAttr)

      Do
        If mvEnv.Connection.IsUnicodeField(pCC.MainAttribute) Then
          vToken = NextToken(pValue, TokenActionTypes.tatSubstitute, True)
        Else
          vToken = NextToken(pValue, TokenActionTypes.tatSubstitute)
        End If

        If vToken <> "" Then
          If pCC.MainDataType = "D" And UCase(vToken) = "NULL" Then vDelimiter = "'"
          If mvToken2 = "" Then
            If vListSet = "'" Then
              If pCC.MainDataType = "D" And UCase(vToken) <> "NULL" Then
                vListSet = mvConn.SQLLiteral("", CDBField.FieldTypes.cftDate, mvToken1)
              Else
                vListSet = vDelimiter & mvToken1 & vDelimiter
              End If
            Else
              If pCC.MainDataType = "D" And UCase(vToken) <> "NULL" Then
                vListSet = vListSet & "," & mvConn.SQLLiteral("", CDBField.FieldTypes.cftDate, mvToken1)
              Else
                vListSet = vListSet & "," & vDelimiter & mvToken1 & vDelimiter
              End If
            End If
            vListTokens = vListTokens + 1
          Else
            vRangeTokens = vRangeTokens + 1
          End If
        End If
      Loop While vToken <> ""

      If vRangeTokens = 0 And (pCC.ValidationTable = "" Or pCC.MainAttribute = pCC.MainValidationAttribute) Then
        mvSingleWhere = mvSingleWhere & "(" & pSelectTable & "." & vMainAttr 'pCC.MainAttribute
        Select Case vListSet
          Case "'null'", "'NULL'"
            mvSingleWhere = mvSingleWhere & " IS NULL"
          Case "'notnull'", "'NOTNULL'"
            mvSingleWhere = mvSingleWhere & " IS NOT NULL"
          Case Else
            If vListTokens = 1 Then
              If pCC.MainDataType = "C" Then
                If InStr(vListSet, "'") > 0 AndAlso InStr(Mid$(vListSet, 2, Len(vListSet) - 2), "'") > 0 Then
                  vListSet = Replace(Mid$(vListSet, 2, Len(vListSet) - 2), "'", "''")
                  vListSet = "'" & vListSet & "'"
                End If
              End If
              If InStr(vListSet, "*") > 0 Or InStr(vListSet, "?") > 0 Then
                If mvEnv.Connection.IsUnicodeField(vMainAttr) Then
                  mvSingleWhere = mvSingleWhere & mvConn.DBLike(Mid(vListSet, 2, Len(vListSet) - 2), CDBField.FieldTypes.cftUnicode)
                Else
                  mvSingleWhere = mvSingleWhere & mvConn.DBLike(Mid(vListSet, 2, Len(vListSet) - 2))
                End If
              Else
                If mvEnv.Connection.IsUnicodeField(vMainAttr) Then
                  mvSingleWhere = mvSingleWhere & " = N" & vListSet
                Else
                  mvSingleWhere = mvSingleWhere & " = " & vListSet
                End If
              End If
            Else
              If pCC.MainDataType = "C" Then
                vListArray = Split(vListSet, ",")
                vListSet = ""
                For vIndex = 0 To UBound(vListArray)
                  vListItem = vListArray(vIndex)
                  If InStr(Mid$(vListItem, 2, Len(vListItem) - 2), "'") > 0 Then
                    vListItem = Replace(Mid$(vListItem, 2, Len(vListItem) - 2), "'", "''")
                    vListItem = "'" & vListItem & "'"
                  End If
                  If Len(vListSet) > 0 Then vListSet = vListSet & ","
                  vListSet = vListSet & vListItem
                Next
              End If
              mvSingleWhere = mvSingleWhere & " IN (" & vListSet & ")"
            End If
        End Select
        mvSingleWhere = mvSingleWhere & ")"

      Else
        If vRangeTokens = 1 And vListTokens = 0 And (pCC.ValidationTable = "" Or pCC.MainAttribute = pCC.MainValidationAttribute) Then
          mvSingleWhere = mvSingleWhere & "(" & pSelectTable & "." & vMainAttr
          If pCC.MainDataType = "D" Then
            mvSingleWhere = mvSingleWhere & mvConn.SQLLiteral("BETWEEN", CDBField.FieldTypes.cftDate, mvToken1) & mvConn.SQLLiteral("AND", CDBField.FieldTypes.cftDate, mvToken2) & ")"
          Else
            mvSingleWhere = mvSingleWhere & " BETWEEN " & vDelimiter & mvToken1 & vDelimiter & " AND " & vDelimiter & mvToken2 & vDelimiter & ")"
          End If
        Else
          If pCC.ValidationTable <> "" Then
            mvSingleWhere = mvSingleWhere & "(" & pSelectTable & "." & vMainAttr
            'ab 15nov01 include testing for nulls and not nulls
            Select Case vListSet
              Case "'null'", "'NULL'"
                mvSingleWhere = mvSingleWhere & " IS NULL"
              Case "'notnull'", "'NOTNULL'"
                mvSingleWhere = mvSingleWhere & " IS NOT NULL"
              Case Else
                mvSingleWhere = mvSingleWhere & " IN ("
                If BuildMainList(pCC, pValue, vDelimiter) Then
                  If vListTokens > 0 And pCC.MainAttribute = pCC.MainValidationAttribute Then
                    mvSingleWhere = mvSingleWhere & "," & vListSet
                  End If
                Else
                  If vListTokens > 0 Then
                    mvSingleWhere = mvSingleWhere & vListSet
                  Else
                    RaiseError(DataAccessErrors.daeNoContactsFound)
                  End If
                End If
                mvSingleWhere = mvSingleWhere & ")"
            End Select
          Else
            mvSingleWhere = mvSingleWhere & "("
            If vListTokens > 0 Then
              mvSingleWhere = mvSingleWhere & pSelectTable & "." & vMainAttr
              mvSingleWhere = mvSingleWhere & " IN ("
              mvSingleWhere = mvSingleWhere & vListSet
              mvSingleWhere = mvSingleWhere & ")"
              vRangeTokens = 1
            Else
              vRangeTokens = 0
            End If
            mvPosition = 1
            Do
              If mvEnv.Connection.IsUnicodeField(pCC.MainAttribute) Then
                vToken = NextToken(pValue, TokenActionTypes.tatSubstitute, True)
              Else
                vToken = NextToken(pValue, TokenActionTypes.tatSubstitute)
              End If
              If vToken <> "" Then
                If mvToken2 <> "" Then
                  vRangeTokens = vRangeTokens + 1
                  If vRangeTokens > 1 Then
                    mvSingleWhere = mvSingleWhere & " or "
                  End If
                  mvSingleWhere = mvSingleWhere & "(" & pSelectTable & "." & vMainAttr
                  If pCC.MainDataType = "D" Then
                    mvSingleWhere = mvSingleWhere & mvConn.SQLLiteral("BETWEEN", CDBField.FieldTypes.cftDate, mvToken1) & mvConn.SQLLiteral("AND", CDBField.FieldTypes.cftDate, mvToken2) & ")"
                  Else
                    mvSingleWhere = mvSingleWhere & " BETWEEN " & vDelimiter & mvToken1 & vDelimiter & " AND " & vDelimiter & mvToken2 & vDelimiter & ")"
                  End If
                End If
              End If
            Loop While vToken <> ""
          End If
          mvSingleWhere = mvSingleWhere & ")"
        End If
      End If
    End Sub
    Private Sub ProcessSameTable(ByRef pCC As CriteriaContext, ByRef pAlias As String, ByRef pTable As String, ByRef pSearchArea As String, ByRef pAttribute As String, ByRef pValue As String)
      Dim vCC As CriteriaContext
      Dim vIndex As Integer

      For Each vCC In mvCriteriaContexts
        If vCC Is pCC Then Exit For 'Find it in the array
      Next vCC
      For vIndex = pCC.ID + 1 To mvCriteriaContexts.Count()
        vCC = CType(mvCriteriaContexts.Item(vIndex), CriteriaContext)
        If pSearchArea <> vCC.SearchArea And pTable = vCC.TableName And Not vCC.Processed Then
          ' possibility to pull together queries on same table
          If (pAttribute <> vCC.MainAttribute And (vCC.Include Or (mvMasterTable = vCC.TableName And vCC.NullsAllowed = False))) Then
            mvSingleWhere = mvSingleWhere & " AND "
            If vCC.Include = False Then mvSingleWhere = mvSingleWhere & "NOT "
            BuildAttrCriteria(vCC, pAlias)
            vCC.Processed = True
          End If
          If pAttribute = vCC.MainAttribute And pValue = vCC.MainValue Then
            mvSingleWhere = mvSingleWhere & " AND "
            If vCC.Include = False Then mvSingleWhere = mvSingleWhere & "NOT "
            If vCC.SubValue <> "" Then ProcessSubsidiaryValue(vCC, (vCC.MainValue), (vCC.SubValue), pAlias)
            If vCC.Period <> "" Then
              If vCC.SubValue <> "" Then mvSingleWhere = mvSingleWhere & " AND "
              ProcessDates(vCC, pAlias)
            End If
            vCC.Processed = True
          End If
        End If
      Next
    End Sub
    Private Sub ProcessSubsidiaryValue(ByRef pCC As CriteriaContext, ByRef pMainValue As String, ByRef pSubsidiaryValue As String, ByRef pSelectTable As String)
      Dim vListTokens As Integer
      Dim vListSet As String
      Dim vRangeTokens As Integer
      Dim vToken As String
      Dim vSubAttr As String

      mvSingleWhere = mvSingleWhere & " ("
      mvPosition = 1
      mvToken1 = ""
      mvToken2 = ""
      vListSet = "'"
      vListTokens = 0
      vRangeTokens = 0
      vSubAttr = pCC.SubAttribute
      If mvConn.IsSpecialColumn(vSubAttr) Then vSubAttr = mvConn.DBSpecialCol("", vSubAttr)
      Do
        If mvEnv.Connection.IsUnicodeField(pCC.SubAttribute) Then
          vToken = NextToken(pSubsidiaryValue, TokenActionTypes.tatSubstitute, True)
        Else
          vToken = NextToken(pSubsidiaryValue, TokenActionTypes.tatSubstitute)
        End If
        If vToken <> "" Then
          If mvToken2 = "" Then
            vListTokens = vListTokens + 1
            If vListSet = "'" Then
              If pCC.SubDataType = "D" And UCase(vToken) <> "NULL" Then
                vListSet = mvConn.SQLLiteral("", CDBField.FieldTypes.cftDate, mvToken1)
              Else
                vListSet = vListSet & mvToken1 & "'"
              End If
            Else
              If pCC.SubDataType = "D" And UCase(vToken) <> "NULL" Then
                vListSet = vListSet & "," & mvConn.SQLLiteral("", CDBField.FieldTypes.cftDate, mvToken1)
              Else
                vListSet = vListSet & ",'" & mvToken1 & "'"
              End If
            End If
          Else
            vRangeTokens = vRangeTokens + 1
          End If
        End If
      Loop While vToken <> ""

      If vRangeTokens = 0 Then
        mvSingleWhere = mvSingleWhere & pSelectTable & "." & vSubAttr
        Select Case vListSet
          Case "'null'", "'NULL'"
            mvSingleWhere = mvSingleWhere & " IS NULL"
          Case "'notnull'", "'NOTNULL'"
            mvSingleWhere = mvSingleWhere & " IS NOT NULL"
          Case Else
            If vListTokens = 1 Then
              If InStr(vListSet, "*") > 0 Or InStr(vListSet, "?") > 0 Then
                If mvEnv.Connection.IsUnicodeField(vSubAttr) Then
                  mvSingleWhere = mvSingleWhere & mvConn.DBLike(Mid(vListSet, 2, Len(vListSet) - 2), CDBField.FieldTypes.cftUnicode)
                Else
                  mvSingleWhere = mvSingleWhere & mvConn.DBLike(Mid(vListSet, 2, Len(vListSet) - 2))
                End If
              Else
                mvSingleWhere = mvSingleWhere & " = " & vListSet
              End If
            Else
              mvSingleWhere = mvSingleWhere & " IN (" & vListSet & ")"
            End If
        End Select
      Else
        If vRangeTokens = 1 And vListTokens = 0 Then
          mvSingleWhere = mvSingleWhere & pSelectTable & "." & vSubAttr
          If pCC.SubDataType = "D" Then
            mvSingleWhere = mvSingleWhere & mvConn.SQLLiteral("BETWEEN", CDBField.FieldTypes.cftDate, mvToken1) & mvConn.SQLLiteral("AND", CDBField.FieldTypes.cftDate, mvToken2)
          Else
            mvSingleWhere = mvSingleWhere & " BETWEEN '" & mvToken1 & "' AND '" & mvToken2 & "'"
          End If
        Else
          If pCC.SubValidationAttribute <> "" Then
            mvSingleWhere = mvSingleWhere & pSelectTable & "." & vSubAttr & " IN ("
            If BuildSubsidiaryList(pCC, pMainValue, pSubsidiaryValue) Then
              If vListTokens > 0 Then mvSingleWhere = mvSingleWhere & "," & vListSet
              mvSingleWhere = mvSingleWhere & ")"
            Else
              If vListTokens > 0 Then
                mvSingleWhere = mvSingleWhere & vListSet
                mvSingleWhere = mvSingleWhere & ")"
              Else
                RaiseError(DataAccessErrors.daeNoContactsFound)
              End If
            End If
          Else
            mvSingleWhere = mvSingleWhere & "("
            If vListTokens > 0 Then
              mvSingleWhere = mvSingleWhere & pSelectTable & "." & vSubAttr
              mvSingleWhere = mvSingleWhere & " IN ("
              mvSingleWhere = mvSingleWhere & vListSet
              mvSingleWhere = mvSingleWhere & ")"
              vRangeTokens = 1
            Else
              vRangeTokens = 0
            End If
            mvPosition = 1
            Do
              If mvEnv.Connection.IsUnicodeField(pCC.SubAttribute) Then
                vToken = NextToken(pSubsidiaryValue, TokenActionTypes.tatSubstitute, True)
              Else
                vToken = NextToken(pSubsidiaryValue, TokenActionTypes.tatSubstitute)
              End If
              If vToken <> "" Then
                If mvToken2 <> "" Then
                  vRangeTokens = vRangeTokens + 1
                  If vRangeTokens > 1 Then
                    mvSingleWhere = mvSingleWhere & " or "
                  End If
                  mvSingleWhere = mvSingleWhere & "(" & pSelectTable & "." & vSubAttr
                  mvSingleWhere = mvSingleWhere & " BETWEEN '" & mvToken1 & "' AND '" & mvToken2 & "')"
                End If
              End If
            Loop While vToken <> ""
            mvSingleWhere = mvSingleWhere & ")"
          End If
        End If
      End If
      mvSingleWhere = mvSingleWhere & ")"
    End Sub
    Public Function RoughCount(ByRef pCriteriaSet As Integer) As Integer
      'This function should produce a rough count of the Records matching the criteria
      Dim vCC As CriteriaContext
      Dim vMinCount As Integer
      Dim vCount As Integer
      Dim vContactCount As Integer
      Dim vRecordCount As Integer
      Dim vInclude As Boolean

      GetCriteriaDetails(pCriteriaSet, True)
      vMinCount = 99999999
      For Each vCC In mvCriteriaContexts
        'Process the records
        vInclude = vCC.Include 'Remember Include
        If mvIndexNoOptimise = False And (vCC.Indexed = False Or vCC.Special = True) Then
          If vCC.Include Then
            vRecordCount = mvConn.GetCount((vCC.TableName), Nothing, "")
            vCount = -vRecordCount
          Else
            If vContactCount = 0 Then vContactCount = mvConn.GetCount(mvMasterTable, Nothing, "")
            vCount = -vContactCount
          End If
        Else
          ClearTableAliases()
          If vCC.Special = True Or vCC.Contacts = False Or (vCC.ValidationTable <> "" And vCC.MainAttribute <> vCC.MainValidationAttribute) Or vCC.SearchArea = "role" Then
            If vCC.Include = False And (vCC.ValidationTable <> "" And vCC.TableName = mvMasterTable And vCC.MainAttribute <> vCC.MainValidationAttribute) Then
              vCC.Include = True
              vCount = SQLCount(vCC)
              vCC.Include = vInclude
              If vContactCount = 0 Then vContactCount = mvConn.GetCount(mvMasterTable, Nothing, "")
              vCount = vContactCount - vCount
            Else
              vCC.Include = True
              vCount = SQLCount(vCC)
              vCC.Include = vInclude
            End If
          Else
            vCC.Include = True
            vCount = SQLCount(vCC)
            vCC.Include = vInclude
            If vCC.Include = False Then
              If vContactCount = 0 Then vContactCount = mvConn.GetCount(mvMasterTable, Nothing, "")
              vCount = vContactCount - vCount
            End If
          End If
        End If

        If vCount = 0 Then
          vMinCount = 0
          Exit For
        Else
          If vCount > 0 Then
            If vMinCount > 0 Then
              If vCount < vMinCount Then vMinCount = vCount
            Else
              If vCount < System.Math.Abs(vMinCount) Then vMinCount = -vCount
            End If
          Else
            If vMinCount > 0 Then
              If System.Math.Abs(vCount) < vMinCount Then
                vMinCount = vCount
              Else
                vMinCount = -vMinCount
              End If
            Else
              If vCount > vMinCount Then vMinCount = vCount
            End If
          End If
        End If
        'Store the count for the criteria line
        vCC.Counted = vCount
        vCC.Save(CriteriaContext.SaveTypes.stUpdateCounted)
      Next vCC
      Count = System.Math.Abs(vMinCount)
      RoughCount = vMinCount
    End Function
    Private Sub SkipToToken(ByRef pString As String, ByRef pLength As Integer)
      Dim vChar As String

      mvPosition = mvPosition + 1
      vChar = Mid(pString, mvPosition, 1)
      While (vChar = " " Or vChar = ",") And mvPosition < pLength
        mvPosition = mvPosition + 1
        vChar = Mid(pString, mvPosition, 1)
      End While
    End Sub
    Private Sub SpecialFirst(ByRef pCC As CriteriaContext)
      Dim vPrevTable1 As String
      Dim vPrevTable2 As String
      Dim vTable1 As String
      Dim vTable2 As String = ""
      Dim vSQL As String
      Dim vSCDTable1 As String
      Dim vSCDTable2 As String
      Dim vSCDAttribute1 As String
      Dim vSCDAttribute2 As String
      Dim vSCDJoinCondition As String
      Dim vRecordSet As CDBRecordSet

      vSQL = "SELECT * FROM selection_control_details WHERE application_name = '" & mvApplication & "'"
      vSQL = vSQL & " AND search_area = '" & pCC.SearchArea & "' AND c_o = '" & pCC.ContactOrOrgValue & "'"
      vSQL = vSQL & " ORDER BY sequence_number"
      vRecordSet = mvConn.GetRecordSet(vSQL)
      If vRecordSet.Fetch() = False Then
        RaiseError(DataAccessErrors.daeExpectedDataMissing)
      Else
        mvSingleWhere = mvSingleWhere & " AND ("
        vPrevTable1 = pCC.TableName
        vPrevTable2 = ""
        vTable1 = "table" & CStr(mvTableNumber)
        Do
          vSCDTable1 = vRecordSet.Fields("table_1").Value
          vSCDTable2 = vRecordSet.Fields("table_2").Value
          vSCDAttribute1 = Replace(vRecordSet.Fields("attribute_1").Value, Chr(34), "")
          If mvConn.IsSpecialColumn(vSCDAttribute1) Then vSCDAttribute1 = mvConn.DBSpecialCol("", vSCDAttribute1)
          vSCDAttribute2 = Replace(vRecordSet.Fields("attribute_2").Value, Chr(34), "")
          If mvConn.IsSpecialColumn(vSCDAttribute2) Then vSCDAttribute2 = mvConn.DBSpecialCol("", vSCDAttribute2)

          vSCDJoinCondition = " " & Replace(vRecordSet.Fields("join_condition").Value, Chr(34), "'") & " "
          If vSCDTable1 <> vPrevTable1 Then
            If vPrevTable2 <> "" Then
              If vSCDTable1 = vPrevTable2 Then
                vTable1 = vTable2
                vTable2 = ""
              Else
                RaiseError(DataAccessErrors.daeInvalidJoinSequence)
              End If
            Else
              RaiseError(DataAccessErrors.daeInvalidJoinSequence)
            End If
          End If
          mvSingleWhere = mvSingleWhere & vTable1 & "." & vSCDAttribute1 & vSCDJoinCondition
          If vSCDAttribute2 <> "" Then
            If vPrevTable2 = "" Or vPrevTable2 <> vSCDTable2 Then
              mvTableNumber = mvTableNumber + 1
              vTable2 = "table" & CStr(mvTableNumber)
              mvSingleTableList = mvSingleTableList & ", " & vSCDTable2 & " " & vTable2
              If Not mvTableAliases.Exists(vTable2) Then mvTableAliases.Add(vSCDTable2, vTable2)
              'BR17963 - changed second part of the condition so that it gets set correctly
              If vSCDTable2 = mvMasterTable AndAlso mvMasterTableAlias <> vTable2 Then mvMasterTableAlias = vTable2
            End If
            mvSingleWhere = mvSingleWhere & vTable2 & "." & vSCDAttribute2
          End If
          vPrevTable1 = vSCDTable1
          vPrevTable2 = vSCDTable2
          vRecordSet.Fetch()
          If vRecordSet.Status() = True Then mvSingleWhere = mvSingleWhere & " AND "
        Loop While vRecordSet.Status() = True
        mvSingleWhere = mvSingleWhere & " )"
      End If
      vRecordSet.CloseRecordSet()
    End Sub
    Private Sub SpecialLater(ByRef pCC As CriteriaContext)
      Dim vPrevTable1 As String
      Dim vPrevTable2 As String
      Dim vTable1 As String
      Dim vTable2 As String
      Dim vSQL As String
      Dim vSCDTable1 As String
      Dim vSCDTable2 As String
      Dim vSCDAttribute1 As String
      Dim vSCDAttribute2 As String
      Dim vSCDJoinCondition As String
      Dim vRecordSet As CDBRecordSet

      vSQL = "SELECT * FROM selection_control_details WHERE application_name = '" & mvApplication & "'"
      vSQL = vSQL & " AND search_area = '" & pCC.SearchArea & "' AND c_o = '" & pCC.ContactOrOrgValue & "'"
      vSQL = vSQL & " ORDER BY sequence_number DESC"
      vRecordSet = mvConn.GetRecordSet(vSQL)
      If vRecordSet.Fetch() = False Then
        RaiseError(DataAccessErrors.daeExpectedDataMissing)
      Else
        vSCDTable1 = vRecordSet.Fields("table_1").Value
        vSCDTable2 = vRecordSet.Fields("table_2").Value

        mvSingleWhere = mvSingleWhere & " ("
        vPrevTable2 = vSCDTable2
        vPrevTable1 = vSCDTable1
        vTable2 = ""
        If pCC.Include Then mvSingleTableList = mvSingleTableList & ", "
        If vSCDTable2 <> "" Then
          vTable2 = "table" & CStr(mvTableNumber)
          mvSingleTableList = mvSingleTableList & vSCDTable2 & " " & vTable2
          If Not mvTableAliases.Exists(vTable2) Then mvTableAliases.Add(vSCDTable2, vTable2)
          mvSingleTableList = mvSingleTableList & ", "
          mvTableNumber = mvTableNumber + 1
        End If
        vTable1 = "table" & CStr(mvTableNumber)
        mvSingleTableList = mvSingleTableList & vSCDTable1 & " " & vTable1
        If Not mvTableAliases.Exists(vTable1) Then mvTableAliases.Add(vSCDTable1, vTable1)
        Do
          vSCDTable1 = vRecordSet.Fields("table_1").Value
          vSCDTable2 = vRecordSet.Fields("table_2").Value
          vSCDAttribute1 = Replace(vRecordSet.Fields("attribute_1").Value, Chr(34), "")
          If mvConn.IsSpecialColumn(vSCDAttribute1) Then vSCDAttribute1 = mvConn.DBSpecialCol("", vSCDAttribute1)
          vSCDAttribute2 = Replace(vRecordSet.Fields("attribute_2").Value, Chr(34), "")
          If mvConn.IsSpecialColumn(vSCDAttribute2) Then vSCDAttribute2 = mvConn.DBSpecialCol("", vSCDAttribute2)
          vSCDJoinCondition = " " & Replace(vRecordSet.Fields("join_condition").Value, Chr(34), "'") & " "
          If vSCDTable2 = "" Then
            mvSingleWhere = mvSingleWhere & vTable1 & "." & vSCDAttribute1 & vSCDJoinCondition
          Else
            If vSCDTable2 = vPrevTable1 Then
              vTable2 = vTable1
              vTable1 = ""
            Else
              If vPrevTable2 <> "" Then
                If vSCDTable2 <> vPrevTable2 Then RaiseError(DataAccessErrors.daeInvalidJoinSequence)
              Else
                RaiseError(DataAccessErrors.daeInvalidJoinSequence)
              End If
            End If
            If vSCDTable1 <> vPrevTable1 Then
              mvTableNumber = mvTableNumber + 1
              vTable1 = "table" & CStr(mvTableNumber)
              mvSingleTableList = mvSingleTableList & ", " & vSCDTable1 & " " & vTable1
              If Not mvTableAliases.Exists(vTable1) Then mvTableAliases.Add(vSCDTable1, vTable1)
            End If
            mvSingleWhere = mvSingleWhere & vTable2 & "." & vSCDAttribute2 & vSCDJoinCondition & vTable1 & "." & vSCDAttribute1
          End If
          vPrevTable1 = vSCDTable1
          vPrevTable2 = vSCDTable2
          vRecordSet.Fetch()
          If vRecordSet.Status() = True Then mvSingleWhere = mvSingleWhere & " AND "
        Loop While vRecordSet.Status() = True
        mvSingleWhere = mvSingleWhere & " )"
      End If
      vRecordSet.CloseRecordSet()
    End Sub
    Private Function SQLCount(ByRef pCC As CriteriaContext) As Integer

      mvTableNumber = 1
      mvLastCTable = ""
      mvLastOTable = ""
      mvCGeographic = ""
      mvOGeographic = ""
      mvCAddressLink = ""
      mvOAddressLink = ""
      mvContactTableAlias = ""
      mvOrgTableAlias = ""
      mvLastCSearchArea = ""
      mvLastOSearchArea = ""
      mvTableList = " "
      mvWhere = " "
      BuildStatementBody(pCC, False)
      LinkToPositions()
      SQLCount = mvConn.GetCount(mvTableList, Nothing, Mid(mvWhere, 8)) 'Step over the WHERE
    End Function
    Private Sub StoreAddressLinks(ByRef pCC As CriteriaContext, ByRef pSelectTable As String)

      If pCC.AddressLink Then
        If pCC.Contacts Then
          If pCC.Geographic Then
            If mvCGeographic = "" Then mvCGeographic = pSelectTable
          Else
            If mvCAddressLink = "" Then mvCAddressLink = pSelectTable
          End If
        Else
          If pCC.Geographic Then
            If mvOGeographic = "" Then mvOGeographic = pSelectTable
          Else
            If mvOAddressLink = "" Then mvOAddressLink = pSelectTable
          End If
        End If
      End If
    End Sub
    Public Sub ClearVariableCriteria()
      mvVariableCriteria = Nothing
      mvVariableCriteria = New Collection
    End Sub
    Public ReadOnly Property ContactCriteriaCount() As Integer
      Get
        ContactCriteriaCount = mvConCriteriaCount
      End Get
    End Property
    Public ReadOnly Property CurrentCriteria() As CriteriaDetails
      Get
        If mvCurrentCriteria Is Nothing Then
          mvCurrentCriteria = New CriteriaDetails
          mvCurrentCriteria.Init(mvEnv)
        End If
        CurrentCriteria = mvCurrentCriteria
      End Get
    End Property
    Public ReadOnly Property OrganisationCriteriaCount() As Integer
      Get
        OrganisationCriteriaCount = mvOrgCriteriaCount
      End Get
    End Property
    Public ReadOnly Property VariableCriteria() As Collection
      Get
        VariableCriteria = mvVariableCriteria
      End Get
    End Property
    Public ReadOnly Property MailingType() As MailingTypes
      Get
        MailingType = mvMailingType
      End Get
    End Property
    Public ReadOnly Property MailingTypeCode() As String
      Get
        If mvMailingType = MailingTypes.mtyStandardExclusions Then
          MailingTypeCode = "GM"
        Else
          MailingTypeCode = mvApplication
        End If
      End Get
    End Property
    Public ReadOnly Property SelectionCount(ByVal pSelectionSet As Integer, ByVal pRevision As Integer) As Integer
      Get
        'Return the number of selected records
        Dim vWhereFields As New CDBFields

        vWhereFields.Add("selection_set", CDBField.FieldTypes.cftLong, pSelectionSet)
        vWhereFields.Add("revision", CDBField.FieldTypes.cftLong, pRevision)
        SelectionCount = mvConn.GetCount(mvSelectionTable, vWhereFields)
      End Get
    End Property
    Public ReadOnly Property Caption() As String
      Get
        Caption = mvCaption
      End Get
    End Property
    Public ReadOnly Property SelectionTable() As String
      Get
        SelectionTable = mvSelectionTable
      End Get
    End Property
    Public ReadOnly Property SelectionSetTable() As String
      Get
        SelectionSetTable = mvSelectionSetTable
      End Get
    End Property
    Public ReadOnly Property MasterAttribute() As String
      Get
        MasterAttribute = mvMasterAttribute
      End Get
    End Property
    Public ReadOnly Property CardProductionType() As MembershipCardProductionTypes
      Get
        CardProductionType = mvCardProductionType
      End Get
    End Property
    Public Property OrganisationRoles() As String
      Get
        OrganisationRoles = mvOrgRoles
      End Get
      Set(ByVal Value As String)
        mvOrgRoles = Value
        mvOrgMailOptionsChanged = True
      End Set
    End Property
    Public Property OrganisationMailTo() As OrgSelectContact
      Get
        OrganisationMailTo = mvOrgMailTo
      End Get
      Set(ByVal Value As OrgSelectContact)
        mvOrgMailTo = Value
        mvOrgMailOptionsChanged = True
      End Set
    End Property
    Public Property OrganisationMailWhere() As OrgSelectAddress
      Get
        OrganisationMailWhere = mvOrgMailWhere
      End Get
      Set(ByVal Value As OrgSelectAddress)
        mvOrgMailWhere = Value
        mvOrgMailOptionsChanged = True
      End Set
    End Property
    Public Property OrganisationAddressUsage() As String
      Get
        OrganisationAddressUsage = mvOrgAddUsage
      End Get
      Set(ByVal Value As String)
        mvOrgAddUsage = Value
        mvOrgMailOptionsChanged = True
      End Set
    End Property
    Public Property OrganisationLabelName() As String
      Get
        OrganisationLabelName = mvOrgLabelName
      End Get
      Set(ByVal Value As String)
        mvOrgLabelName = Value
        mvOrgMailOptionsChanged = True
      End Set
    End Property
    Public Property VariableParameters() As String
      Get
        Dim vVarCriteria As VariableCriteria

        mvVariableParameters = ""
        For Each vVarCriteria In mvVariableCriteria
          If mvVariableParameters.Length > 0 Then mvVariableParameters = mvVariableParameters & "|"
          mvVariableParameters = mvVariableParameters & vVarCriteria.VariableName & "=" & vVarCriteria.Value
        Next vVarCriteria
        mvVariableParameters = "|" & mvVariableParameters & "|"

        VariableParameters = mvVariableParameters
      End Get
      Set(ByVal Value As String)
        mvVariableParameters = Value
      End Set
    End Property

    Public ReadOnly Property SummaryPrintValid() As Boolean
      Get
        Select Case mvMailingType
          Case MailingTypes.mtyGeneralMailing, MailingTypes.mtyDirectDebits, MailingTypes.mtyStandingOrders, MailingTypes.mtyStandingOrderCancellation, MailingTypes.mtyMembers, MailingTypes.mtyMembershipCards, MailingTypes.mtyPayers, MailingTypes.mtySubscriptions, MailingTypes.mtyIrishGiftAid
            SummaryPrintValid = True
        End Select
      End Get
    End Property

    Public ReadOnly Property TempTableName() As String
      Get
        Dim vTableName As String

        If Mid(Right(mvTempTableName, 2), 1, 1) = "_" Then
          vTableName = Left(mvTempTableName, Len(mvTempTableName) - 2)
        Else
          vTableName = mvTempTableName
        End If
        TempTableName = vTableName
      End Get
    End Property

    Public ReadOnly Property ListTableName() As String
      Get
        If mvAppealMailing Then
          ListTableName = TempTableName
        Else
          ListTableName = mvSelectionTable
        End If
      End Get
    End Property

    Public ReadOnly Property AppealMailing() As Boolean
      Get
        AppealMailing = mvAppealMailing
      End Get
    End Property

    Public ReadOnly Property DisplayOrgSelection() As Boolean
      Get
        DisplayOrgSelection = mvDisplayOrgSelection
      End Get
    End Property

    Public Property IncludeHistoricRoles() As Boolean
      Get
        IncludeHistoricRoles = mvIncludeHistoricRoles
      End Get
      Set(ByVal Value As Boolean)
        mvIncludeHistoricRoles = Value
      End Set
    End Property

    Public Property OrganisationMailOptionsChanged() As Boolean
      Get
        OrganisationMailOptionsChanged = mvOrgMailOptionsChanged
      End Get
      Set(ByVal Value As Boolean)
        mvOrgMailOptionsChanged = Value
      End Set
    End Property
    Public Property Segment() As Segment
      Get
        Segment = mvSegment
      End Get
      Set(ByVal Value As Segment)
        mvSegment = Value
      End Set
    End Property

    Public ReadOnly Property SegmentsHaveSelectionOptions() As Boolean
      Get
        SegmentsHaveSelectionOptions = mvSegmentsHaveSelectionOptions
      End Get
    End Property

    Public ReadOnly Property ContactAttribute() As String
      Get
        ContactAttribute = mvContactAttribute
      End Get
    End Property
    Public Sub DeleteCriteriaDetails(ByVal pCriteriaSet As Integer)
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("criteria_set", CDBField.FieldTypes.cftLong, pCriteriaSet)
      mvConn.DeleteRecords("criteria_set_details", vWhereFields, False)
    End Sub
    Public Sub SetOrgMailWhere(ByVal pOrgMailWhere As String)
      Select Case pOrgMailWhere
        Case "O"
          mvOrgMailWhere = OrgSelectAddress.osaOrganisationAddress
        Case "D"
          mvOrgMailWhere = OrgSelectAddress.osaDefaultAddress
        Case "U"
          mvOrgMailWhere = OrgSelectAddress.osaAddressByUsage
        Case Else
          mvOrgMailWhere = OrgSelectAddress.osaNone
      End Select
      mvOrgMailOptionsChanged = True
    End Sub
    Public Sub SetOrgMailTo(ByVal pOrgMailTo As String)
      Select Case pOrgMailTo
        Case "A"
          mvOrgMailTo = OrgSelectContact.oscAllEmployees
        Case "D"
          mvOrgMailTo = OrgSelectContact.oscDefaultContact
        Case "O"
          mvOrgMailTo = OrgSelectContact.oscOrganisation
        Case Else
          mvOrgMailTo = OrgSelectContact.oscNone
      End Select
      mvOrgMailOptionsChanged = True
    End Sub
    Public Function GetOrgMailWhere() As String
      Dim vOrgMailWhere As String

      Select Case mvOrgMailWhere
        Case OrgSelectAddress.osaAddressByUsage
          vOrgMailWhere = "U"
        Case OrgSelectAddress.osaDefaultAddress
          vOrgMailWhere = "D"
        Case OrgSelectAddress.osaOrganisationAddress
          vOrgMailWhere = "O"
        Case Else
          vOrgMailWhere = ""
      End Select

      GetOrgMailWhere = vOrgMailWhere
    End Function
    Public Function GetOrgMailTo() As String
      Dim vOrgMailTo As String

      Select Case mvOrgMailTo
        Case OrgSelectContact.oscAllEmployees
          vOrgMailTo = "A"
        Case OrgSelectContact.oscDefaultContact
          vOrgMailTo = "D"
        Case OrgSelectContact.oscOrganisation
          vOrgMailTo = "O"
        Case Else
          vOrgMailTo = ""
      End Select

      GetOrgMailTo = vOrgMailTo
    End Function

    Public Sub ProcessSegments(ByRef pAppeal As Appeal, ByRef pJob As JobSchedule)
      Dim vSegment As Segment
      Dim vCC As CriteriaContext
      Dim vCriteriaSets As Collection = Nothing
      Dim vDedupedTableName As String
      Dim vStatuses As String = ""
      Dim vMsg As String
      Dim vCount As Integer
      Dim vTable As String
      Dim vStatusArray() As String
      Dim vIndex As Integer
      Dim vSSNo As Integer
      Dim vCriteriaSet As CriteriaSet

      vMsg = XLAT("Generating Mailing Data - ")
      pJob.InfoMessage = vMsg & XLAT("Processing Segments")

      mvAppealMailing = Not pAppeal.Campaign = "SMCam"
      mvSegmentsHaveSelectionOptions = pAppeal.SegmentSelectionOptions <> CDBEnvironment.SelectionOptionSettings.sosNone
      If mvAppealMailing Then
        mvVariableParameters = pAppeal.VariableParameters
        If pAppeal.SegmentSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosNone Then
          SetOrgMailTo(pAppeal.OrgMailTo)
          SetOrgMailWhere(pAppeal.OrgMailWhere)
          mvOrgRoles = pAppeal.OrgMailRoles
          mvOrgAddUsage = pAppeal.OrgMailAddrUsage
          mvOrgLabelName = pAppeal.OrgMailLabelName
        End If
      End If

      vStatusArray = Split(pAppeal.JointStatusExclusions, ",")
      For vIndex = 0 To UBound(vStatusArray)
        vStatuses = vStatuses & "'" & vStatusArray(vIndex) & "',"
      Next
      If Len(vStatuses) > 0 Then vStatuses = Left(vStatuses, Len(vStatuses) - 1)

      'create temporary tables
      If Not mvAppealMailing Then vSSNo = CType(pAppeal.Segments.Item(1), Segment).SelectionSet
      GenerateTempTableName((pAppeal.Campaign), (pAppeal.AppealCode), vSSNo)
      vDedupedTableName = MailingTableName((pAppeal.Campaign), (pAppeal.AppealCode))
      CreateTempTables(False, vDedupedTableName, mvTempTableName, False, True, (pAppeal.MailJoints))

      'Add a segment zero for the processing of the standard exclusion criteria
      vSegment = New Segment
      vSegment.Init(mvEnv, (pAppeal.Campaign), (pAppeal.AppealCode), "Exclud")
      vSegment.CriteriaSet = ExclusionCriteriaSet
      If vSegment.CriteriaSet > 0 Then pAppeal.Segments.Add(vSegment, "", 1)
      'Write SQL Log Header
      mvConn.LogMailSQL("***** User : " & mvEnv.User.Logname & ", Date : " & Now.ToString & ", Application : " & mvCaption & " *****")
      'Will any of the segments Score or Randomise the selected records?
      For Each vSegment In pAppeal.Segments
        If vSegment.Score.Length > 0 Or vSegment.Random Then
          mvSegmentScoreOrRandom = True
          Exit For
        End If
      Next vSegment
      'Process each segment in appeal
      For Each vSegment In pAppeal.Segments
        vCount = vCount + 1
        If mvSegmentScoreOrRandom Then
          vTable = If(vCount = 1, mvTempTableName, vDedupedTableName)
        Else
          vTable = mvTempTableName
        End If
        If vSegment.SegmentSequence = 0 Then
          pJob.InfoMessage = vMsg & XLAT("Processing Standard Exclusions")
        Else
          pJob.InfoMessage = vMsg & XLATP1("Processing Segment '%s'", vSegment.SegmentDesc)
        End If
        If vSegment.SegmentSequence > 0 And vSegment.SelectionSet = 0 Then
          vSegment.SelectionSet = mvEnv.GetControlNumber("SS")
        End If
        ClearTableAliases(True)
        vSegment.ActualCount = 0

        If pAppeal.SegmentSelectionOptions <> CDBEnvironment.SelectionOptionSettings.sosNone Then
          If vSegment.SelectionOptions <> CDBEnvironment.SelectionOptionSettings.sosNone Then
            SetOrgMailTo(vSegment.OrgMailTo)
            SetOrgMailWhere(vSegment.OrgMailWhere)
            mvOrgRoles = vSegment.OrgMailRoles
            mvOrgAddUsage = vSegment.OrgMailAddrUsage
            mvOrgLabelName = vSegment.OrgMailLabelName
          ElseIf pAppeal.SelectionOptions <> CDBEnvironment.SelectionOptionSettings.sosNone Then
            SetOrgMailTo(pAppeal.OrgMailTo)
            SetOrgMailWhere(pAppeal.OrgMailWhere)
            mvOrgRoles = pAppeal.OrgMailRoles
            mvOrgAddUsage = pAppeal.OrgMailAddrUsage
            mvOrgLabelName = pAppeal.OrgMailLabelName
          End If
        End If

        If vSegment.CriteriaSet > 0 Then
          vCriteriaSet = New CriteriaSet
          vCriteriaSet.Init(mvEnv, vSegment.CriteriaSet)
          If vCriteriaSet.SelectionSteps.Count() = 0 Then
            If Len(mvVariableParameters) > 0 And mvVariableParameters <> "||" Then
              InitVariableCriteria(vSegment.CriteriaSet)
            Else
              GetCriteriaDetails(vSegment.CriteriaSet, True)
            End If
            For Each vCC In mvCriteriaContexts
              If vSegment.SegmentSequence = 0 Then
                If Not vCC.Include Then
                  vCC.Include = True
                  vCC.AndOr = "or"
                End If
              End If
              If pAppeal.BypassCount Then
                vCC.Counted = vCC.SequenceNumber
              Else
                If vCC.Include Then
                  vCC.Counted = SQLCount(vCC)
                Else
                  vCC.Counted = mvConn.GetCount((vCC.TableName), Nothing, "")
                End If
                'reset table aliases collection
                ClearTableAliases()
              End If
              vCC.Save(CriteriaContext.SaveTypes.stUpdateCounted)
            Next vCC
            'break up criteria set into a collection of one or more criteria sets
            If mvCriteriaContexts.Count() > 0 Then
              vCriteriaSets = CriteriaPreprocessor()
            Else
              'BR13818: Raise error if both criteria set and details are not found and critera set number is provided. Previously we were raising this error in GetCriteriaDetails but the code was commented for BR2340.
              If vCriteriaSet.Existing = False Then RaiseError(DataAccessErrors.daeNoCriteria)
              vCriteriaSets = New Collection
            End If
          End If
          'select records
          SelectContacts(vCriteriaSet, vCriteriaSets, pAppeal, vSegment, vTable)
        Else
          vCriteriaSets = New Collection
          vCriteriaSet = New CriteriaSet
          vCriteriaSet.Init(mvEnv)
          SelectContacts(vCriteriaSet, vCriteriaSets, pAppeal, vSegment, vTable)
        End If
        If mvSegmentScoreOrRandom Then ConvertAndDedupSegments(pAppeal, pJob, vMsg, vStatuses, vDedupedTableName, vCount, vSegment)
      Next vSegment
      If mvSegmentScoreOrRandom Then
        'remove all records from dedup table where selection_set is zero
        mvConn.ExecuteSQL("DELETE FROM " & vDedupedTableName & " WHERE selection_set = 0")
      Else
        ConvertAndDedupSegments(pAppeal, pJob, vMsg, vStatuses, vDedupedTableName)
      End If
      'Write SQL Log Footer
      mvConn.LogMailSQL("***** Date : " & Now.ToString)
    End Sub

    Public Function MailingTableName(Optional ByRef pCampaignCode As String = "", Optional ByRef pAppealCode As String = "") As String
      If Not mvAppealMailing Then
        MailingTableName = mvSelectionTable
      Else
        MailingTableName = "ca_" & LCase(pCampaignCode) & "_" & LCase(pAppealCode)
      End If
    End Function

    Public Function GenerateTempTableName(ByRef pCampaignCode As String, ByRef pAppealCode As String, ByRef pSelectionSet As Integer) As String
      mvTempTableName = "_tmp_" & LCase(pCampaignCode) & "_" & LCase(pAppealCode)
      If Not mvAppealMailing Then mvTempTableName = mvTempTableName & "_" & pSelectionSet 'pAppeal.Segments(1).SelectionSet
      mvTempTableName = Mid(mvEnv.User.Logname, 1, 28 - Len(mvTempTableName)) & mvTempTableName
      GenerateTempTableName = mvTempTableName
    End Function

    Private Function CriteriaPreprocessor() As Collection
      Dim vCC As CriteriaContext
      Dim vRecordSet As CDBRecordSet
      Dim vLevel As Integer
      Dim vOldLevel As Integer
      Dim vCriteriaSet As Integer
      Dim vLeftLen As Integer
      Dim vRightLen As Integer
      Dim vOldSet As Integer
      Dim vTempCS As String
      Dim vTempCC As CriteriaContext
      Dim vSQL As String = ""
      Dim vORSets As New Collection
      Dim vORSetsStack As New Collection
      Dim vCurrentSets As New Collection
      Dim vCurrentSetsStack As New Collection
      Dim vLevelSets As New Collection
      Dim vLevelSetStack As New Collection
      Dim vCriteriaSets As New Collection
      Dim vBracketStack As New Collection
      Dim vSequenceStack As New Collection
      Dim vBracket As String
      Dim vSequence As Integer

      vCriteriaSet = mvEnv.GetControlNumber("CS")
      vCurrentSets.Add(vCriteriaSet)
      vCriteriaSets.Add(vCriteriaSet)
      vLevel = 1
      vOldLevel = 1

      For Each vCC In mvCriteriaContexts
        If vCC.ID = 1 Then
          If vCC.AndOr.Length > 0 Then vCC.AndOr = ""
        End If
        'clear out opening & closing brackets on a single item
        vLeftLen = Len(vCC.LeftParenthesis)
        vRightLen = Len(vCC.RightParenthesis)
        If vLeftLen > 0 And vRightLen > 0 Then
          If vLeftLen = vRightLen Then
            vCC.LeftParenthesis = ""
            vCC.RightParenthesis = ""
          ElseIf vLeftLen > vRightLen Then
            vCC.RightParenthesis = ""
            vCC.LeftParenthesis = Mid(vCC.LeftParenthesis, vRightLen + 1)
          Else
            vCC.LeftParenthesis = ""
            vCC.RightParenthesis = Mid(vCC.RightParenthesis, vLeftLen + 1)
          End If
        End If
        If vCC.AndOr = "or" Then 'check for an OR
          If vCC.LeftParenthesis.Length > 0 Then
            vLevel = vLevel + Len(vCC.LeftParenthesis)
            While vOldLevel < vLevel
              AddSetToStack(vORSetsStack, vORSets)
              vORSets = New Collection
              vBracketStack.Add("or")
              vLevel = vLevel - 1
            End While
          End If
          'copy all criteria in upper levels to create base for new criteria sets
          If vLevelSetStack.Count() > 0 Then
            vLevelSets = New Collection
            vLevelSets = CType(vLevelSetStack.Item(vLevelSetStack.Count()), Collection)
            For Each vTempCS In vLevelSets
              If vSQL.Length > 0 Then vSQL = vSQL & ", "
              vSQL = vSQL & vTempCS
            Next vTempCS
          End If
          If vSQL.Length > 0 Then
            vOldSet = 0
            vCurrentSets = New Collection
            If vSequenceStack.Count() > 0 Then
              vSequence = CInt(vSequenceStack.Item(vSequenceStack.Count()))
            Else
              vSequence = 0
            End If
            vRecordSet = mvConn.GetRecordSet("SELECT * FROM criteria_set_details WHERE criteria_set IN (" & vSQL & ") AND sequence_number <= " & vSequence & " ORDER BY criteria_set, sequence_number")
            While vRecordSet.Fetch() = True
              If vOldSet <> vRecordSet.Fields("criteria_set").IntegerValue Then
                vCriteriaSet = mvEnv.GetControlNumber("CS")
                vOldSet = vRecordSet.Fields("criteria_set").IntegerValue
                'add new criteria set to full & current list
                vCriteriaSets.Add(vCriteriaSet)
                vCurrentSets.Add(vCriteriaSet)
                AddToAllSetNosInStack(vCurrentSetsStack, vCriteriaSet)
                vORSets.Add(vCriteriaSet)
                AddToAllSetNosInStack(vORSetsStack, vCriteriaSet)
              End If
              vRecordSet.Fields("criteria_set").Value = CStr(vCriteriaSet)
              mvConn.InsertRecord("criteria_set_details", vRecordSet.Fields)
            End While
            vRecordSet.CloseRecordSet()
            If vCurrentSets.Count() = 0 Then 'No sequence numbers to copy from upper level - create an empty set
              vCriteriaSet = mvEnv.GetControlNumber("CS")
              vCriteriaSets.Add(vCriteriaSet)
              vCurrentSets.Add(vCriteriaSet)
              AddToAllSetNosInStack(vCurrentSetsStack, vCriteriaSet)
              vORSets.Add(vCriteriaSet)
              AddToAllSetNosInStack(vORSetsStack, vCriteriaSet)
            End If
          Else 'No upper levels to copy - create an empty set
            vCriteriaSet = mvEnv.GetControlNumber("CS")
            vCriteriaSets.Add(vCriteriaSet)
            vCurrentSets = New Collection
            vCurrentSets.Add(vCriteriaSet)
            AddToAllSetNosInStack(vCurrentSetsStack, vCriteriaSet)
            vORSets.Add(vCriteriaSet)
            AddToAllSetNosInStack(vORSetsStack, vCriteriaSet)
          End If
          'reset all relevant stacks if at top level because everything that has gone before is irrelevant
          If vLevel = 1 Then
            vCurrentSetsStack = New Collection
            vSequenceStack = New Collection
            vLevelSetStack = New Collection
          End If
        Else 'its an actual or implicit and - so check for a bracket
          If vCC.LeftParenthesis.Length > 0 Then
            vLevel = vLevel + Len(vCC.LeftParenthesis)
            While vOldLevel < vLevel
              AddSetToStack(vORSetsStack, vORSets)
              vORSets = New Collection
              vBracketStack.Add("and")
              vSequenceStack.Add((vCC.SequenceNumber - 1))
              AddSetToStack(vLevelSetStack, vCurrentSets)
              AddSetToStack(vCurrentSetsStack, vCurrentSets)
              vOldLevel = vOldLevel + 1
            End While
          End If
        End If
        'add criteria record to all relevant criteria sets
        For Each vTempCS In vCurrentSets
          vTempCC = New CriteriaContext
          vTempCC.Clone(mvEnv, mvConn, vCC)
          vTempCC.CriteriaSet = IntegerValue(vTempCS)
          vTempCC.AndOr = ""
          vTempCC.LeftParenthesis = ""
          vTempCC.RightParenthesis = ""
          vTempCC.Save(CriteriaContext.SaveTypes.stInsert)
        Next vTempCS
        'handle right parentheses
        If vCC.RightParenthesis.Length > 0 Then
          vLevel = vLevel - Len(vCC.RightParenthesis)
          While vLevel < vOldLevel
            If vBracketStack.Count() > 0 Then
              vBracket = CStr(vBracketStack.Item(vBracketStack.Count()))
              vBracketStack.Remove(vBracketStack.Count())
            Else
              vBracket = ""
            End If
            If vBracket = "or" Then
              vLevel = vLevel + 1
              vCurrentSets = New Collection
              For Each vTempCS In vORSets
                vCurrentSets.Add(vTempCS)
              Next vTempCS
              vORSets = New Collection
              vORSets = CType(vORSetsStack.Item(vORSetsStack.Count()), Collection)
              vORSetsStack.Remove(vORSetsStack.Count())
            Else
              vOldLevel = vOldLevel - 1
              vORSets = New Collection
              If vORSetsStack.Count() > 0 Then
                vORSets = CType(vORSetsStack.Item(vORSetsStack.Count()), Collection)
                vORSetsStack.Remove(vORSetsStack.Count())
              End If
              If vSequenceStack.Count() > 0 Then vSequenceStack.Remove(vSequenceStack.Count()) '????
              If vLevelSetStack.Count() > 0 Then vLevelSetStack.Remove(vLevelSetStack.Count())
              vCurrentSets = New Collection
              If vCurrentSetsStack.Count() > 0 Then
                vCurrentSets = CType(vCurrentSetsStack.Item(vCurrentSetsStack.Count()), Collection)
                vCurrentSetsStack.Remove(vCurrentSetsStack.Count())
              End If
            End If
          End While
        End If
        vCC.Save(CriteriaContext.SaveTypes.stUpdateBrackets)
      Next vCC
      CriteriaPreprocessor = vCriteriaSets
    End Function

    Private Sub AddToAllSetNosInStack(ByRef pStackColl As Collection, ByRef pCriteriaSetNo As Integer)
      Dim vTempColl As New Collection

      'Add the new criteria set number into each collection within the stack collection
      For Each vTempColl In pStackColl
        vTempColl.Add(pCriteriaSetNo)
      Next vTempColl
    End Sub
    Public Sub CreateTempTables(ByVal pCreateDedupOnly As Boolean, ByVal pDedupTableName As String, Optional ByVal pWorkTableName As String = "", Optional ByVal pCreateWorkOnly As Boolean = False, Optional ByVal pCreateWork2 As Boolean = False, Optional ByVal pMailJoints As Boolean = False)
      Dim vWhereFields As New CDBFields
      With vWhereFields
        .Add("selection_set", CDBField.FieldTypes.cftLong)
        .Add("revision", CDBField.FieldTypes.cftInteger)
        If mvMasterAttribute <> mvContactAttribute Then .Add(mvMasterAttribute, CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("address_number_2", CDBField.FieldTypes.cftLong)
        If mvAppealMailing Then vWhereFields.Add("segment_sequence", CDBField.FieldTypes.cftLong)
        If mvMailingType = MailingTypes.mtyIrishGiftAid Then .Add("performance", CDBField.FieldTypes.cftCharacter, "6")
        vWhereFields.Add("marker", CDBField.FieldTypes.cftCharacter, "1")
      End With
      If Not pCreateDedupOnly And Len(pWorkTableName) > 0 Then 'Build work table
        mvConn.CreateTableFromFields(pWorkTableName, vWhereFields)
      End If
      If Not pCreateWorkOnly Then 'Create dedup table(s), if it doesn't exist already
        DeleteSelection(SelectionSetNumber, Revision, True)
        mvConn.CreateTableFromFields(pDedupTableName, vWhereFields)
        If pMailJoints And mvMasterAttribute <> mvContactAttribute Then mvConn.CreateTableFromFields(pDedupTableName & "_1", vWhereFields)
      End If
      If pMailJoints And Not pCreateDedupOnly And Len(pWorkTableName) > 0 Then
        mvConn.CreateTableFromFields(pWorkTableName & "_4", vWhereFields)
        With vWhereFields
          .Add("joint_contact_number", CDBField.FieldTypes.cftLong)
          .Add("joint_address_number", CDBField.FieldTypes.cftLong)
        End With
        mvConn.CreateTableFromFields(pWorkTableName & "_3", vWhereFields)
        mvConn.CreateTableFromFields(pWorkTableName & "_6", vWhereFields)
      End If
      If pCreateWork2 Then
        vWhereFields = New CDBFields
        With vWhereFields
          If mvAppealMailing Then
            .Add("segment_sequence", CDBField.FieldTypes.cftLong)
          Else
            .Add("selection_set", CDBField.FieldTypes.cftLong)
          End If
          If mvMasterAttribute <> mvContactAttribute Then
            .Add(mvMasterAttribute, CDBField.FieldTypes.cftLong)
          Else
            .Add("contact_number", CDBField.FieldTypes.cftLong)
          End If
          mvConn.CreateTableFromFields(pWorkTableName & "_2", vWhereFields)
          If pMailJoints And mvMasterAttribute <> mvContactAttribute Then
            .Remove(2)
            .Add("contact_number", CDBField.FieldTypes.cftLong)
            mvConn.CreateTableFromFields(pWorkTableName & "_5", vWhereFields)
          End If
        End With
      End If
    End Sub

    Public Sub CreateExamCertificateTempTable(pTableName As String, pExamUnitLinkId As Integer, pCertRunType As String)
      If mvEnv.Connection.TableExists(pTableName & "_esuh") Then
        mvEnv.Connection.DropTable(pTableName & "_esuh")
      End If
      If mvEnv.Connection.TableExists(pTableName & "_orig") Then
        If mvEnv.Connection.TableExists(pTableName) Then
          mvEnv.Connection.DropTable(pTableName)
        End If
        If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
          Call New SQLStatement(mvEnv.Connection, "EXEC sp_rename '" & pTableName & "_orig', '" & pTableName & "'").GetIntegerValue()
        Else
          Call New SQLStatement(mvEnv.Connection, "ALTER TABLE " & pTableName & "_orig RENAME TO " & pTableName).GetIntegerValue()
        End If
      End If

      Dim vSql As New StringBuilder

      vSql.AppendLine("WITH cte(exam_unit_link_id, parent_unit_link_id) ")
      vSql.AppendLine("     AS (SELECT eul.exam_unit_link_id, ")
      vSql.AppendLine("                eul.parent_unit_link_id ")
      vSql.AppendLine("         FROM   exam_unit_links eul ")
      vSql.AppendLine("         WHERE  eul.exam_unit_link_id = " & pExamUnitLinkId.ToString & " ")
      vSql.AppendLine("         UNION ALL ")
      vSql.AppendLine("         SELECT eul.exam_unit_link_id, ")
      vSql.AppendLine("                eul.parent_unit_link_id ")
      vSql.AppendLine("         FROM   exam_unit_links eul ")
      vSql.AppendLine("                inner join cte ")
      vSql.AppendLine("                        ON cte.exam_unit_link_id = eul.parent_unit_link_id) ")
      vSql.AppendLine("SELECT cte.exam_unit_link_id ")
      vSql.AppendLine("FROM   cte")
      vSql.AppendLine("       INNER JOIN exam_unit_cert_run_types eucrt")
      vSql.AppendLine("         ON eucrt.exam_unit_link_id = cte.exam_unit_link_id ")
      vSql.AppendLine("            AND eucrt.exam_cert_run_type = '" & pCertRunType & "'")
      Dim vRequiredUnits As IEnumerable(Of Integer) = From vRow As DataRow In New SQLStatement(mvEnv.Connection, vSql.ToString).GetDataTable().AsEnumerable
                                                      Select CInt(vRow("exam_unit_link_id"))
      vSql.Length = 0

      Dim vFirst As Boolean = True
      If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then
        vSql.AppendLine("CREATE TABLE " & pTableName & "_esuh")
        vSql.AppendLine("AS")
      End If
      For Each vLinkId As Integer In vRequiredUnits
        If Not vFirst Then
          vSql.AppendLine("UNION ALL ")
        End If
        vSql.AppendLine("(SELECT sc.selection_set, ")
        vSql.AppendLine("       sc.revision, ")
        vSql.AppendLine("       sc.exam_student_unit_header_id, ")
        vSql.AppendLine("       sc.contact_number ")
        If vFirst And mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
          vSql.AppendLine("INTO " & pTableName & "_esuh")
        End If
        vSql.AppendLine("FROM   " & pTableName & " sc ")
        vSql.AppendLine("       INNER JOIN exam_student_unit_header esuh ")
        vSql.AppendLine("               ON esuh.exam_student_unit_header_id = sc.exam_student_unit_header_id ")
        vSql.AppendLine("       INNER JOIN dbo.exam_unit_links eul ")
        vSql.AppendLine("               ON eul.exam_unit_link_id = esuh.exam_unit_link_id ")
        mvExamUnitCertRunType = ExamUnitCertRunType.GetInstance(mvEnv, vLinkId, ExamCertRunType.GetInstance(mvEnv, pCertRunType))
        If mvExamUnitCertRunType IsNot Nothing Then
          If mvExamUnitCertRunType.IncludeView.Length > 0 Then
            vSql.AppendLine("       INNER JOIN " & mvExamUnitCertRunType.IncludeView & " inc ")
            vSql.AppendLine("               ON inc.exam_student_unit_header_id = sc.exam_student_unit_header_id ")
          End If
          If mvExamUnitCertRunType.ExcludeView.Length > 0 Then
            vSql.AppendLine("WHERE  Coalesce(eul.base_unit_link_id,eul.exam_unit_link_id) = " & vLinkId)
            If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then
              vSql.AppendLine("MINUS ")
            Else
              vSql.AppendLine("EXCEPT ")
            End If
            vSql.AppendLine("SELECT sc.selection_set, ")
            vSql.AppendLine("       sc.revision, ")
            vSql.AppendLine("       sc.exam_student_unit_header_id, ")
            vSql.AppendLine("       sc.contact_number ")
            vSql.AppendLine("FROM   " & pTableName & " sc ")
            vSql.AppendLine("       INNER JOIN exam_student_unit_header esuh ")
            vSql.AppendLine("               ON esuh.exam_student_unit_header_id = sc.exam_student_unit_header_id ")
            vSql.AppendLine("       INNER JOIN dbo.exam_unit_links eul ")
            vSql.AppendLine("               ON eul.exam_unit_link_id = esuh.exam_unit_link_id ")
            vSql.AppendLine("       INNER JOIN " & mvExamUnitCertRunType.ExcludeView & " inc ")
            vSql.AppendLine("               ON inc.exam_student_unit_header_id = sc.exam_student_unit_header_id ")
          End If
        End If
        vSql.AppendLine("WHERE  Coalesce(eul.base_unit_link_id,eul.exam_unit_link_id) = " & vLinkId & ")")
        vFirst = False
      Next vLinkId

      Call New SQLStatement(mvEnv.Connection, vSql.ToString).GetIntegerValue()

      vSql.Length = 0
      If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
        vSql.AppendLine("EXEC sp_rename '" & pTableName & "', '" & pTableName & "_orig'")
      Else
        vSql.AppendLine("ALTER TABLE " & pTableName & " RENAME TO " & pTableName & "_orig")
      End If
      Call New SQLStatement(mvEnv.Connection, vSql.ToString).GetIntegerValue()

      vSql.Length = 0
      If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
        vSql.AppendLine("EXEC sp_rename '" & pTableName & "_esuh', '" & pTableName & "'")
      Else
        vSql.AppendLine("ALTER TABLE " & pTableName & "_esuh RENAME TO " & pTableName)
      End If
      Call New SQLStatement(mvEnv.Connection, vSql.ToString).GetIntegerValue()
    End Sub

    Public Sub PostProcessCertificates(vFilename As String)
      Dim tempFileName As String = Path.GetTempPath & Path.GetRandomFileName
      Using rawData As New CsvReader(vFilename)
        Using processedData As New StreamWriter(tempFileName)
          Using dataProcessor As New CertificateDataProcessor(mvEnv,
                                                              rawData,
                                                              mvExamUnitCertRunType) With {.StreamWriter = processedData}
            dataProcessor.Process()
          End Using
        End Using
      End Using
      File.Delete(vFilename)
      File.Move(tempFileName, vFilename)
    End Sub

    Public Sub ResetExamCertificateTempTable(pTableName As String)
      If mvEnv.Connection.TableExists(pTableName & "_eul") Then
        mvEnv.Connection.DropTable(pTableName & "_eul")
      End If
      If mvEnv.Connection.TableExists(pTableName & "_orig") Then
        If mvEnv.Connection.TableExists(pTableName) Then
          mvEnv.Connection.DropTable(pTableName)
        End If
        If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
          Call New SQLStatement(mvEnv.Connection, "EXEC sp_rename '" & pTableName & "_orig', '" & pTableName & "'").GetIntegerValue()
        Else
          Call New SQLStatement(mvEnv.Connection, "ALTER TABLE " & pTableName & "_orig RENAME TO " & pTableName).GetIntegerValue()
        End If
      End If
    End Sub

    Private Sub DedupRecords(ByVal pDedupTable As String, ByVal pWorkTable As String, Optional ByVal pMailJoints As Boolean = False)
      Dim vUsageOrDefault As Boolean
      Dim vAttrList As String
      Dim vSQL As String
      Dim vControlAttr As String
      Dim vUniqueID As String
      Dim vWhere As String
      Dim vWorkTable As String

      If Mid(Right(pWorkTable, 2), 1, 1) = "_" Then
        vWorkTable = Left(pWorkTable, Len(pWorkTable) - 2)
      Else
        vWorkTable = pWorkTable
      End If

      If pMailJoints And mvMasterAttribute <> mvContactAttribute Then
        'Create list of unique contact numbers
        If mvAppealMailing Then
          vControlAttr = "segment_sequence"
        Else
          vControlAttr = "selection_set"
        End If
        vUniqueID = "contact_number"
        vSQL = "INSERT INTO " & vWorkTable & "_5 (" & vControlAttr & ",contact_number)"
        vSQL = vSQL & " SELECT MIN(" & vControlAttr & "),contact_number FROM " & pWorkTable
        vSQL = vSQL & " GROUP BY " & vUniqueID
        If Not mvSegmentScoreOrRandom Then vSQL = vSQL & " HAVING MIN(" & vControlAttr & ") > 0"
        mvConn.LogMailSQL(vSQL)
        mvConn.ExecuteSQL(vSQL)
        mvConn.CreateIndex(True, vWorkTable & "_5", {vControlAttr, "contact_number"})

        'Dedup records in temporary table using the unique list of contact numbers
        vUsageOrDefault = (mvOrgMailWhere = OrgSelectAddress.osaAddressByUsage Or mvOrgMailWhere = OrgSelectAddress.osaDefaultAddress)
        vAttrList = "revision,y.contact_number"
        vAttrList = "y.selection_set," & vAttrList
        vSQL = "INSERT INTO " & pDedupTable & "_1 (selection_set,"
        vSQL = vSQL & "revision,contact_number,"
        vSQL = vSQL & "segment_sequence,"
        If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ","
        vSQL = vSQL & "address_number,address_number_2)"
        vSQL = vSQL & " SELECT " & vAttrList & ",MIN(y.segment_sequence)"
        If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & ", MIN(" & mvMasterAttribute & ")"
        If vUsageOrDefault Then
          vSQL = vSQL & ",MIN(a.address_number)"
        Else
          vSQL = vSQL & ",MIN(y.address_number)"
        End If
        vSQL = vSQL & ",MIN(y.address_number)" 'to populate the address_number_2 attribute
        vSQL = vSQL & " FROM " & vWorkTable & "_5 x, " & pWorkTable & " y"
        vWhere = " WHERE x." & vControlAttr & " = y." & vControlAttr
        vWhere = vWhere & " AND x.contact_number = y.contact_number"
        If vUsageOrDefault Then
          If mvOrgMailWhere = OrgSelectAddress.osaDefaultAddress Then vSQL = vSQL & ", contacts c"
          If mvOrgMailWhere = OrgSelectAddress.osaAddressByUsage Then vSQL = vSQL & ", contact_address_usages cau"
          vSQL = vSQL & ", addresses a"
          'vSQL = vSQL & ", addresses a, organisation_addresses oa, organisations o, contact_positions cp"
          vSQL = vSQL & vWhere
          If mvOrgMailWhere = OrgSelectAddress.osaDefaultAddress Then
            vSQL = vSQL & " AND y.contact_number = c.contact_number AND c.address_number = a.address_number"
          ElseIf mvOrgMailWhere = OrgSelectAddress.osaAddressByUsage Then
            vSQL = vSQL & " AND y.contact_number = cau.contact_number AND address_usage = '" & mvOrgAddUsage & "' AND cau.address_number = a.address_number"
          End If
          'SDT/MD 1/5/2001
          'The following lines will cause contacts whose default or usage address that
          'is not an organisation address to be omitted. So they will be commented out for now
          'also see FROM clause above
          '   vSQL = vSQL & " AND a.address_number = oa.address_number AND oa.organisation_number = o.organisation_number"
          '   vSQL = vSQL & " AND o.organisation_number = cp.organisation_number AND cp.contact_number = "
          '    If mvOrgMailWhere = osaDefaultAddress Then
          '      vSQL = vSQL & "c.contact_number"
          '    Else
          '      vSQL = vSQL & "y.contact_number"
          '    End If
        Else
          vSQL = vSQL & vWhere
        End If
        vSQL = vSQL & " GROUP BY " & vAttrList
        mvConn.LogMailSQL(vSQL)
        mvConn.ExecuteSQL(vSQL)
        If mvMasterAttribute <> mvContactAttribute Then
          mvConn.CreateIndex(False, pDedupTable & "_1", {"selection_set", mvMasterAttribute, "contact_number", "address_number"})
        Else
          mvConn.CreateIndex(False, pDedupTable & "_1", {"selection_set", "contact_number", "address_number"})
        End If
        If mvAppealMailing Then
          mvConn.CreateIndex(False, pDedupTable & "_1", {"segment_sequence", mvMasterAttribute})
        Else
          mvConn.CreateIndex(False, pDedupTable & "_1", {"selection_set", mvMasterAttribute})
        End If
        mvConn.CreateIndex(False, pDedupTable & "_1", {mvMasterAttribute})
      End If

      'Create list of unique master attributes
      If mvAppealMailing Then
        vControlAttr = "segment_sequence"
      Else
        vControlAttr = "selection_set"
      End If
      If mvMasterAttribute <> mvContactAttribute Then
        vUniqueID = mvMasterAttribute
      Else
        vUniqueID = "contact_number"
      End If
      vSQL = "INSERT INTO " & vWorkTable & "_2 (" & vControlAttr & "," & vUniqueID & ")"
      vSQL = vSQL & " SELECT MIN(" & vControlAttr & ")," & vUniqueID & " FROM " '& pWorkTable
      If pMailJoints And mvMasterAttribute <> mvContactAttribute Then
        vSQL = vSQL & pDedupTable & "_1"
      Else
        vSQL = vSQL & pWorkTable
      End If
      vSQL = vSQL & " GROUP BY " & vUniqueID
      If Not mvSegmentScoreOrRandom Then vSQL = vSQL & " HAVING MIN(" & vControlAttr & ") > 0"
      mvConn.LogMailSQL(vSQL)
      mvConn.ExecuteSQL(vSQL)
      mvConn.CreateIndex(True, vWorkTable & "_2", {vControlAttr, vUniqueID})

      'Dedup records in temporary table using the unique list of master attributes
      vUsageOrDefault = (mvOrgMailWhere = OrgSelectAddress.osaAddressByUsage Or mvOrgMailWhere = OrgSelectAddress.osaDefaultAddress)
      vAttrList = "revision"
      If pMailJoints And mvMasterAttribute <> mvContactAttribute Then
        vAttrList = vAttrList & ", y." & mvMasterAttribute
        If mvAppealMailing And mvSegmentScoreOrRandom Then vAttrList = vAttrList & ",y.segment_sequence"
      Else
        If mvAppealMailing And mvSegmentScoreOrRandom Then vAttrList = vAttrList & ",y.segment_sequence"
        vAttrList = vAttrList & ", y.contact_number"
      End If
      If mvMasterAttribute <> mvContactAttribute And Not pMailJoints Then vAttrList = "x." & mvMasterAttribute & "," & vAttrList
      vAttrList = "y.selection_set," & vAttrList
      If mvMailingType = MailingTypes.mtyIrishGiftAid Then vAttrList = vAttrList & ", performance"
      vSQL = "INSERT INTO " & pDedupTable & " (selection_set,"
      If mvMasterAttribute <> mvContactAttribute And Not pMailJoints Then vSQL = vSQL & mvMasterAttribute & ","
      vSQL = vSQL & "revision,"
      If pMailJoints And mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ","
      If mvAppealMailing And mvSegmentScoreOrRandom Then vSQL = vSQL & "segment_sequence,"
      vSQL = vSQL & "contact_number,"
      If mvMailingType = MailingTypes.mtyIrishGiftAid Then vSQL = vSQL & "performance,"
      vSQL = vSQL & "address_number,address_number_2"
      vSQL = vSQL & ")"
      vSQL = vSQL & " SELECT " & vAttrList
      If pMailJoints And mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & ", MIN(y.contact_number)"
      If vUsageOrDefault Then
        vSQL = vSQL & ",MIN(a.address_number)"
      Else
        vSQL = vSQL & ",MIN(y.address_number)"
      End If
      vSQL = vSQL & ",MIN(y.address_number)" 'to populate the address_number_2 attribute
      vSQL = vSQL & " FROM " & vWorkTable & "_2 x, "
      If pMailJoints And mvMasterAttribute <> mvContactAttribute Then
        vSQL = vSQL & pDedupTable & "_1"
      Else
        vSQL = vSQL & pWorkTable
      End If
      vSQL = vSQL & " y"
      vWhere = " WHERE x." & vControlAttr & " = y." & vControlAttr
      vWhere = vWhere & " AND x." & vUniqueID & " = y." & vUniqueID
      If vUsageOrDefault Then
        If mvOrgMailWhere = OrgSelectAddress.osaDefaultAddress Then vSQL = vSQL & ", contacts c"
        If mvOrgMailWhere = OrgSelectAddress.osaAddressByUsage Then vSQL = vSQL & ", contact_address_usages cau"
        vSQL = vSQL & ", addresses a"
        'vSQL = vSQL & ", addresses a, organisation_addresses oa, organisations o, contact_positions cp"
        vSQL = vSQL & vWhere
        If mvOrgMailWhere = OrgSelectAddress.osaDefaultAddress Then
          vSQL = vSQL & " AND y.contact_number = c.contact_number AND c.address_number = a.address_number"
        ElseIf mvOrgMailWhere = OrgSelectAddress.osaAddressByUsage Then
          vSQL = vSQL & " AND y.contact_number = cau.contact_number AND address_usage = '" & mvOrgAddUsage & "' AND cau.address_number = a.address_number"
        End If
        'SDT/MD 1/5/2001
        'The following lines will cause contacts whose default or usage address that
        'is not an organisation address to be omitted. So they will be commented out for now
        'also see FROM clause above
        '   vSQL = vSQL & " AND a.address_number = oa.address_number AND oa.organisation_number = o.organisation_number"
        '   vSQL = vSQL & " AND o.organisation_number = cp.organisation_number AND cp.contact_number = "
        '    If mvOrgMailWhere = osaDefaultAddress Then
        '      vSQL = vSQL & "c.contact_number"
        '    Else
        '      vSQL = vSQL & "y.contact_number"
        '    End If
      Else
        vSQL = vSQL & vWhere
      End If
      vSQL = vSQL & " GROUP BY " & vAttrList
      mvConn.LogMailSQL(vSQL)
      mvConn.ExecuteSQL(vSQL)

      'Remove list of unique contact numbers & master attributes
      If pMailJoints And mvMasterAttribute <> mvContactAttribute Then
        vSQL = "DROP TABLE " & vWorkTable & "_5"
        mvConn.ExecuteSQL(vSQL)
        vSQL = "DROP TABLE " & pDedupTable & "_1"
        mvConn.ExecuteSQL(vSQL)
      End If
      vSQL = "DROP TABLE " & vWorkTable & "_2"
      mvConn.ExecuteSQL(vSQL)
    End Sub
    Public Sub RecreateSelectionTable(ByVal pSourceTable As String)
      Dim vWhereFields As New CDBFields
      Dim vSQL As String

      'Drop & recreate Selection Table
      If mvConn.TableExists(mvSelectionTable) Then mvConn.DropTable(mvSelectionTable)
      With vWhereFields
        .Add("selection_set", CDBField.FieldTypes.cftLong)
        .Add("revision", CDBField.FieldTypes.cftInteger)
        If mvMasterAttribute <> mvContactAttribute Then .Add(mvMasterAttribute, CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
      End With
      mvConn.CreateTableFromFields(mvSelectionTable, vWhereFields)

      'Populate Selection Table from Source Table
      vSQL = "INSERT INTO " & mvSelectionTable & " SELECT selection_set,revision,"
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ","
      vSQL = vSQL & "contact_number,address_number FROM " & pSourceTable
      mvConn.ExecuteSQL(vSQL)

      'Create index on Selection Table
      mvConn.CreateIndex(True, mvSelectionTable, {"selection_set", "revision", mvMasterAttribute})
    End Sub

    Private Sub ConvertAndDedupSegments(ByRef pAppeal As Appeal, ByRef pJob As JobSchedule, ByVal pMsg As String, ByVal pStatuses As String, ByVal pDedupedTableName As String, Optional ByVal pSegmentCount As Integer = 0, Optional ByRef pSegment As Segment = Nothing)
      Dim vSegment As Segment
      Dim vSQL As String
      Dim vRecordSet As CDBRecordSet
      Dim vScoreOrRandomise As Boolean
      Dim vTempTableName As String = ""
      Dim vSequence As Integer

      If mvSegmentScoreOrRandom And pSegmentCount > 1 Then
        'recreate temp table
        CreateTempTables(False, pDedupedTableName, mvTempTableName, True, True, (pAppeal.MailJoints))
        'copy everything from dedup to temp
        vSQL = "INSERT INTO " & mvTempTableName & " (segment_sequence, selection_set, revision, "
        If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
        vSQL = vSQL & "contact_number, address_number, address_number_2) SELECT segment_sequence, selection_set, revision, "
        If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
        vSQL = vSQL & "contact_number, address_number, address_number_2 FROM " & pDedupedTableName
        mvConn.LogMailSQL(vSQL)
        mvConn.ExecuteSQL(vSQL)
        'drop and recreate dedup table
        mvConn.ExecuteSQL("DROP TABLE " & pDedupedTableName)
        CreateTempTables(True, pDedupedTableName, "", False, False, (pAppeal.MailJoints))
      End If

      Dim vIndexes As New CDBIndexes
      vIndexes.Init(mvConn, mvTempTableName)
      'Create indices on temp table
      If mvMasterAttribute <> mvContactAttribute Then
        vIndexes.CreateIfMissing(mvConn, False, {"selection_set", mvMasterAttribute, "contact_number", "address_number"})
      Else
        vIndexes.CreateIfMissing(mvConn, False, {"selection_set", "contact_number", "address_number"})
      End If
      If mvAppealMailing Then
        vIndexes.CreateIfMissing(mvConn, False, {"segment_sequence", If(pAppeal.MailJoints, "contact_number", mvMasterAttribute)})
      Else
        vIndexes.CreateIfMissing(mvConn, False, {"selection_set", mvMasterAttribute})
      End If
      vIndexes.CreateIfMissing(mvConn, False, {If(pAppeal.MailJoints, "contact_number", mvMasterAttribute)})

      If pAppeal.SegmentSelectionOptions = CDBEnvironment.SelectionOptionSettings.sosNone Then
        If mvOrgMailWhere > 0 And mvOrgMailTo = OrgSelectContact.oscOrganisation And mvOrgRoles <> "" Then
          pJob.InfoMessage = pMsg & XLAT("Processing Roles")
          ProcessRoles(mvTempTableName)
        End If
      Else
        For Each vSegment In pAppeal.Segments
          If vSegment.SelectionOptions <> CDBEnvironment.SelectionOptionSettings.sosNone Then
            SetOrgMailTo(vSegment.OrgMailTo)
            SetOrgMailWhere(vSegment.OrgMailWhere)
            mvOrgRoles = vSegment.OrgMailRoles
            mvOrgAddUsage = vSegment.OrgMailAddrUsage
            mvOrgLabelName = vSegment.OrgMailLabelName
          ElseIf pAppeal.SelectionOptions <> CDBEnvironment.SelectionOptionSettings.sosNone Then
            SetOrgMailTo(pAppeal.OrgMailTo)
            SetOrgMailWhere(pAppeal.OrgMailWhere)
            mvOrgRoles = pAppeal.OrgMailRoles
            mvOrgAddUsage = pAppeal.OrgMailAddrUsage
            mvOrgLabelName = pAppeal.OrgMailLabelName
          End If
          If mvOrgMailWhere > 0 And mvOrgMailTo = OrgSelectContact.oscOrganisation And mvOrgRoles <> "" Then
            pJob.InfoMessage = pMsg & XLAT("Processing Roles for Segment " & vSegment.SegmentDesc)
            ProcessRoles(mvTempTableName, vSegment.SelectionSet)
          End If
        Next vSegment
      End If

      If pAppeal.MailJoints Then
        pJob.InfoMessage = pMsg & XLAT("Converting to Joint Contacts")
        'BR 8900, ConvertToJoints changes mvTempTableName for Dedup process.
        'Store so it can be reliably reset after dedup
        vTempTableName = mvTempTableName
        If mvSegmentScoreOrRandom Then vSequence = pSegment.SegmentSequence
        ConvertToJoints(mvTempTableName, pStatuses, pAppeal, vSequence)
      End If
      'Dedup those records selected - populate dedup table
      pJob.InfoMessage = pMsg & XLAT("Deduplicating Selected Records")
      DedupRecords(pDedupedTableName, mvTempTableName, (pAppeal.MailJoints))
      'Create index on dedup table
      mvConn.CreateIndex(True, pDedupedTableName, {"selection_set", "revision", mvMasterAttribute})
      'Remove joint-conversion work table
      If pAppeal.MailJoints Then
        'BR 8900 Reset Temp Table Name following dedup:
        mvTempTableName = vTempTableName
        mvConn.ExecuteSQL("DROP TABLE " & mvTempTableName & "_3")
        mvConn.ExecuteSQL("DROP TABLE " & mvTempTableName & "_4")
        mvConn.ExecuteSQL("DROP TABLE " & mvTempTableName & "_6")
      End If
      'Remove work table
      mvConn.ExecuteSQL("DROP TABLE " & mvTempTableName)
      If mvAppealMailing Then
        'Remove Segment 0 - standard exclusions segment
        If Not mvSegmentScoreOrRandom Then
          vSegment = New Segment
          vSegment = CType(pAppeal.Segments.Item(1), Access.Segment)
          If vSegment.CriteriaSet > 0 And vSegment.SegmentSequence = 0 Then
            pAppeal.Segments.Remove(1)
          End If
        End If
        'Now count the number of records in each segment
        If mvSegmentScoreOrRandom Then
          pJob.InfoMessage = pMsg & XLAT("Determining Actual Count for Segment " & pSegment.SegmentDesc)
        Else
          pJob.InfoMessage = pMsg & XLAT("Determining Actual Segment Counts")
        End If
        vSQL = "SELECT x.selection_set, count(*)  AS  set_count"
        vSQL = vSQL & " FROM " & pDedupedTableName & " x"
        If mvSegmentScoreOrRandom Then vSQL = vSQL & " WHERE x.selection_set = " & pSegment.SelectionSet
        vSQL = vSQL & " GROUP BY x.selection_set"
        vRecordSet = mvConn.GetRecordSet(vSQL)
        With vRecordSet
          While .Fetch() = True
            For Each vSegment In pAppeal.Segments
              If vSegment.SelectionSet = .Fields(1).IntegerValue Then
                vSegment.ActualCount = .Fields(2).IntegerValue
                Exit For
              End If
            Next vSegment
          End While
          .CloseRecordSet()
        End With
        'Now that the records are deduped and counted do some other things
        If Not mvSegmentScoreOrRandom Or (mvSegmentScoreOrRandom And pSegmentCount = 1) Then pAppeal.ActualCount = CStr(0)
        For Each vSegment In pAppeal.Segments
          If mvSegmentScoreOrRandom Then
            vScoreOrRandomise = False
            If vSegment.SelectionSet = pSegment.SelectionSet And vSegment.SegmentSequence > 0 Then
              vScoreOrRandomise = True
            End If
          Else
            vScoreOrRandomise = True
          End If
          If vScoreOrRandomise Then
            'Process score, if req'd
            If vSegment.Score.Length > 0 And vSegment.ActualCount > vSegment.RequiredCount Then
              pJob.InfoMessage = pMsg & XLATP1("Scoring Segment '%s'", vSegment.SegmentDesc)
              ScoreContacts(vSegment, pDedupedTableName)
            End If
            'Process random, if req'd
            If vSegment.Random And vSegment.ActualCount > vSegment.RequiredCount Then
              pJob.InfoMessage = pMsg & XLATP1("Selecting Random Contacts for Segment '%s'", vSegment.SegmentDesc)
              SelectRandomContacts(vSegment, pDedupedTableName)
            End If
          End If
          If (mvSegmentScoreOrRandom And vScoreOrRandomise) Or Not mvSegmentScoreOrRandom Then
            vSegment.Save()
            pAppeal.ActualCount = CStr(CDbl(pAppeal.ActualCount) + vSegment.ActualCount)
            If mvSegmentScoreOrRandom And vScoreOrRandomise Then Exit For
          End If
        Next vSegment
      End If

      If mvSegmentScoreOrRandom Then
        'Drop index on temp table
        If mvMasterAttribute <> mvContactAttribute Then
          mvConn.DropIndex(mvTempTableName, {"selection_set", mvMasterAttribute, "contact_number", "address_number"})
        Else
          mvConn.DropIndex(mvTempTableName, {"selection_set", "contact_number", "address_number"})
        End If
        If mvAppealMailing Then mvConn.DropIndex(mvTempTableName, {"segment_sequence", mvMasterAttribute})
        'Drop index on dedup table
        mvConn.DropIndex(pDedupedTableName, {"selection_set", "revision", mvMasterAttribute})
      End If
    End Sub
    Public Function DetermineSortOrder(ByVal pSortOrder As SortOrderTypes, ByVal pAdditionalSort As Boolean, ByVal pAdditionalSort2 As Boolean) As String
      Dim vOrder As String = ""

      If MailingType <> MailingTypes.mtySubscriptions Then
        If MailingType = MailingTypes.mtyMembers Then
          pAdditionalSort2 = False
        Else
          pAdditionalSort = False
          pAdditionalSort2 = False
        End If
      End If
      Select Case pSortOrder
        Case SortOrderTypes.sotBranch
          Select Case MailingType
            Case MailingTypes.mtyMembershipCards 'member number
              vOrder = "member_number,co.uk,co.region,a.country,a.postcode,c.surname,c.forenames"
            Case MailingTypes.mtyMembers
              If pAdditionalSort Then
                vOrder = "m.branch,a.address_number,o.order_number,mt.card_order,surname,c.contact_number"
              Else
                vOrder = "m.branch,surname,forenames"
              End If
            Case MailingTypes.mtySubscriptions
              If pAdditionalSort Then
                vOrder = "a.branch,a.address_number,s.order_number,surname,forenames,sc.contact_number,s.despatch_method,s.product"
              ElseIf pAdditionalSort2 Then
                vOrder = "a.branch,a.address_number,s.product,surname,forenames,sc.contact_number,s.despatch_method"
              Else
                vOrder = "a.branch,surname,forenames,sc.contact_number,sc.address_number,s.despatch_method,s.product"
              End If
            Case Else
              vOrder = "a.branch,c.surname,c.forenames"
          End Select

        Case SortOrderTypes.sotCountry
          Select Case MailingType
            Case MailingTypes.mtyMembershipCards 'membership type, gift & member number
              vOrder = "mt.membership_type,co.uk,co.region,a.country,a.postcode,c.surname,c.forenames"
            Case MailingTypes.mtyMembers
              If pAdditionalSort Then
                vOrder = "a.country,town,a.address_number,o.order_number,mt.card_order,surname,c.contact_number"
              Else
                vOrder = "a.country,town,surname"
              End If
            Case MailingTypes.mtySubscriptions
              If pAdditionalSort Then
                vOrder = "a.country,a.address_number,s.order_number,surname,forenames,sc.contact_number,s.despatch_method,s.product"
              ElseIf pAdditionalSort2 Then
                vOrder = "a.country,a.address_number,s.product,surname,forenames,sc.contact_number,s.despatch_method"
              Else
                vOrder = "a.country,surname,forenames,sc.contact_number,sc.address_number,s.despatch_method,s.product"
              End If
            Case Else
              vOrder = "a.country,a.postcode,c.surname,c.forenames"
          End Select

        Case SortOrderTypes.sotMailsort
          Select Case MailingType
            Case MailingTypes.mtyMembershipCards
              vOrder = "direct,a.sortcode,a.postcode,a.address_number,c.surname,c.forenames"
            Case MailingTypes.mtyMembers
              If pAdditionalSort Then
                vOrder = "direct,a.sortcode,postcode,a.address_number,o.order_number,mt.card_order,surname,c.contact_number"
              Else
                vOrder = "direct,a.sortcode,postcode,a.address_number,surname,forenames"
              End If
            Case MailingTypes.mtySubscriptions
              If pAdditionalSort Then
                vOrder = "direct,a.sortcode,a.address_number,s.order_number,surname,forenames,sc.contact_number,s.despatch_method,s.product"
              ElseIf pAdditionalSort2 Then
                vOrder = "direct,a.sortcode,a.address_number,s.product,surname,forenames,sc.contact_number,s.despatch_method"
              Else
                vOrder = "direct,a.sortcode,surname,forenames,sc.contact_number,sc.address_number,s.despatch_method,s.product"
              End If
            Case Else
              vOrder = "direct,a.sortcode,a.postcode,c.surname"
          End Select

        Case SortOrderTypes.sotSurname
          Select Case MailingType
            Case MailingTypes.mtyMembers
              If pAdditionalSort Then
                vOrder = "a.address_number,o.order_number,mt.card_order,surname,c.contact_number"
              Else
                vOrder = "surname,forenames"
              End If
            Case MailingTypes.mtySubscriptions
              If pAdditionalSort Then
                vOrder = "sc.address_number,surname,s.order_number,forenames,sc.contact_number,s.despatch_method,s.product"
              ElseIf pAdditionalSort2 Then
                vOrder = "sc.address_number,surname,forenames,sc.contact_number,s.despatch_method,s.product"
              Else
                vOrder = "surname,s.order_number,forenames,sc.contact_number,sc.address_number,s.despatch_method,s.product"
              End If
            Case Else
              vOrder = "c.surname,c.forenames"
          End Select

        Case SortOrderTypes.sotOther1
          Select Case MailingType
            Case MailingTypes.mtyDirectDebits
              vOrder = "sc.direct_debit_number"
            Case MailingTypes.mtyStandingOrders
              vOrder = "sc.bankers_order_number"
            Case MailingTypes.mtyMemberFulfilment   'BR17283
              vOrder = "sc.order_number"
            Case MailingTypes.mtyMembers
              If pAdditionalSort Then
                vOrder = "a.address_number,o.order_number,m.member_number,mt.card_order,surname,c.contact_number"
              Else
                vOrder = "m.member_number"
              End If
            Case MailingTypes.mtySubscriptions
              If pAdditionalSort Then
                vOrder = "s.despatch_method,sc.address_number,s.order_number,surname,forenames,sc.contact_number,s.product"
              ElseIf pAdditionalSort2 Then
                vOrder = "s.despatch_method,sc.address_number,s.product,surname,forenames,sc.contact_number"
              Else
                vOrder = "s.despatch_method,sc.contact_number,sc.address_number,s.product"
              End If
            Case MailingTypes.mtyMembershipCards  'BR19690
              vOrder = "m.member_number"
          End Select

        Case SortOrderTypes.sotOther2
          Select Case MailingType
            Case MailingTypes.mtyMembers 'membership type for Member mailings
              If pAdditionalSort Then
                vOrder = "mt.membership_type,a.address_number,o.order_number,surname,c.contact_number"
              Else
                vOrder = "mt.membership_type,surname,forenames"
              End If
            Case MailingTypes.mtyMembershipCards 'membership type, gift & member number
              vOrder = "mt.membership_type,o.gift_membership,m.member_number,c.surname,c.forenames"  'BR19690
            Case Else
              '
          End Select
      End Select

      If InStr(vOrder, MasterAttribute) = 0 Then
        'Add master attribute to sort order if it is not already there SuppLog 21066
        If vOrder.Length > 0 Then vOrder = vOrder & ","
        vOrder = vOrder & "sc." & MasterAttribute
      End If
      DetermineSortOrder = vOrder
    End Function

    Private Function TableContainsMaster(ByVal pTableAlias As String) As String
      Dim vTable As String

      On Error GoTo TableContainsMasterError

      If CStr(mvTableAliases(pTableAlias)) <> mvMasterTable Then
        If mvMasterAttrTables.Exists(CStr(mvTableAliases(pTableAlias))) Then
          vTable = pTableAlias
        Else
          vTable = mvMasterTableAlias
        End If
      Else
        vTable = pTableAlias
      End If
      TableContainsMaster = vTable
      Exit Function

TableContainsMasterError:
      'An error may occur due to the supplied table alias not yet existing in the mvTableAliases collection.
      'There are places in the code where a table alias is built and used in the construction of a SQL statement
      'before the code knows what the actual table is, and therefore before the alias has been added to the collection.
      'At this stage, rather than raising an error, just return the supplied table alias...and hope for the best.
      TableContainsMaster = pTableAlias
    End Function

    Public Function CriteriaContainsORs(ByVal pCriteriaSet As Integer) As Boolean
      Dim vCC As CriteriaContext
      Dim vContainsORs As Boolean

      GetCriteriaDetails(pCriteriaSet, True)
      For Each vCC In mvCriteriaContexts
        If vCC.AndOr = "or" Then
          vContainsORs = True
          Exit For
        End If
      Next vCC
      CriteriaContainsORs = vContainsORs
    End Function
    Private Sub AddSetToStack(ByRef pStackColl As Collection, ByRef pSetColl As Collection)
      Dim vTempColl As New Collection
      Dim vTempCS As Object

      'Copy the items from the Set collection into a temporary collection.
      'This is done to ensure that each collection added to the Stack collection is unique
      For Each vTempCS In pSetColl
        vTempColl.Add(vTempCS)
      Next vTempCS
      'Add the temporary collection into the Stack collection
      pStackColl.Add(vTempColl)
    End Sub

    Private Sub ClearTableAliases(Optional ByVal pClearMasterAlias As Boolean = False)
      mvTableAliases.Clear()
      If pClearMasterAlias Then mvMasterTableAlias = ""
    End Sub

    Public Function LMSelectDataSQL(ByRef pViewName As String, ByRef pRestrictions As CDBFields, ByVal pAllRecords As Boolean, ByVal pViewingSelection As Boolean, ByRef pMaxRecords As Integer, Optional ByVal pCountOnly As Boolean = False) As String
      Dim vSQL As String
      Dim vPrefix As String
      Dim vSuffix As String = ""
      Dim vFilter As Boolean

      If pRestrictions.Count > 0 Then
        vFilter = True
        pMaxRecords = FILTERED_RECORD_LIMIT
      Else
        vFilter = False
        If pViewingSelection Then
          pMaxRecords = FILTERED_RECORD_LIMIT
        Else
          pMaxRecords = UNFILTERED_RECORD_LIMIT
        End If
      End If
      If pAllRecords Then pMaxRecords = MAX_RECORDS_IN_GRID

      If (mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer) Then
        vPrefix = "SELECT /* SQLServerCSC */ TOP " & pMaxRecords & " x.* FROM "
      Else
        vPrefix = "SELECT * FROM (SELECT x.*, rownum AS rownumber FROM "          'Changed for BR15474 to restrict by rownumber in an outer select
      End If
      If pViewingSelection Then
        vSQL = ListTableName & " sc LEFT OUTER JOIN " & pViewName & " x ON sc.contact_number = x.contact_number"
        vSQL = vSQL & " WHERE sc.selection_set = " & SelectionSetNumber
        If vFilter Then vSQL = vSQL & " AND " & LMGetWhereClause(pRestrictions, "x")
        vSQL = vSQL & " AND x.contact_number IS NOT NULL"
        If mvConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then
          vSuffix = ") WHERE rownumber < " & pMaxRecords + 1
          'The following was added for BR13599 and then removed for BR13773 - Now superceeded by BR15474
          'Else
          '  vSuffix = mvEnv.Connection.DBForceOrder
        End If
      Else
        vSQL = pViewName & " x "
        vSQL = vSQL & StepOwnershipRestriction()
        If vFilter Then
          vSQL = vSQL & LMGetWhereClause(pRestrictions, "x") & " AND "
        Else
          vSQL = vSQL & LMGetNoFilterRestriction()
        End If
        vSQL = vSQL & " x.contact_number NOT IN ( SELECT contact_number FROM " & ListTableName & " WHERE selection_set = " & SelectionSetNumber & ")"
        If (mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle) Then vSuffix = ") WHERE rownumber < " & pMaxRecords + 1
      End If
      If pCountOnly Then
        Return mvEnv.Connection.ProcessAnsiJoins("SELECT Count(*) FROM " & vSQL)
      Else
        Return mvEnv.Connection.ProcessAnsiJoins(vPrefix & vSQL & vSuffix)
      End If
    End Function

    Public Function LMSelectEbuDataSQL(ByRef pViewName As String, ByRef pRestrictions As CDBFields, ByVal pAllRecords As Boolean, ByVal pViewingSelection As Boolean, ByRef pMaxRecords As Integer, Optional ByVal pCountOnly As Boolean = False) As String
      Dim vSQL As String
      Dim vPrefix As String
      Dim vSuffix As String = ""
      Dim vFilter As Boolean

      If pRestrictions.Count > 0 Then
        vFilter = True
        pMaxRecords = FILTERED_RECORD_LIMIT
      Else
        vFilter = False
        If pViewingSelection Then
          pMaxRecords = FILTERED_RECORD_LIMIT
        Else
          pMaxRecords = UNFILTERED_RECORD_LIMIT
        End If
      End If
      If pAllRecords Then pMaxRecords = MAX_RECORDS_IN_GRID

      If (mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer) Then
        vPrefix = "SELECT /* SQLServerCSC */ TOP " & pMaxRecords & " x.* FROM "
      Else
        vPrefix = "SELECT * FROM (SELECT x.*, rownum AS rownumber FROM "          'Changed for BR15474 to restrict by rownumber in an outer select
      End If
      If pViewingSelection Then
        vSQL = ListTableName & " sc LEFT OUTER JOIN " & pViewName & " x ON sc.exam_booking_unit_id = x.exam_booking_unit_id"
        vSQL = vSQL & " WHERE sc.selection_set = " & SelectionSetNumber
        vSQL = vSQL & " AND x.exam_booking_unit_id IS NOT NULL"
        If mvConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then
          vSuffix = ") WHERE rownumber < " & pMaxRecords + 1
        End If
      Else
        vSQL = pViewName & " x "
        vSQL = vSQL & "WHERE x.exam_booking_unit_id NOT IN ( SELECT exam_booking_unit_id FROM " & ListTableName & " WHERE selection_set = " & SelectionSetNumber & ")"
        If (mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle) Then vSuffix = ") WHERE rownumber < " & pMaxRecords + 1
      End If
      If pCountOnly Then
        Return mvEnv.Connection.ProcessAnsiJoins("SELECT Count(*) FROM " & vSQL)
      Else
        Return mvEnv.Connection.ProcessAnsiJoins(vPrefix & vSQL & vSuffix)
      End If
    End Function

    Public Function LMGetNoFilterRestriction() As String
      If mvEnv.Connection.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then
        'If mvLMMaxContact = 0 Then mvLMMaxContact = IntegerValue(mvEnv.Connection.GetValue("SELECT MAX(contact_number) FROM (SELECT contact_number FROM contacts WHERE contact_type <> 'O' ORDER BY contact_number ) WHERE rownum < 2000"))
        'Return " x.contact_number < " & mvLMMaxContact & " AND "
        Return ""         'Using the new mechanism for rownum it does not seem necessary to add this restriction any more
      Else
        Return ""
      End If
    End Function

    Public Function LMGetWhereClause(ByRef pRestrictions As CDBFields, ByRef pAlias As String) As String
      Dim vWhere As String
      If pRestrictions.ContainsKey("address") Then
        'In some client databases this is a text field in others a varchar. Text field value needs modifying so that
        'it can be easily used in comparison such is LIKE, <> etc... 
        pRestrictions("address").Name = mvConn.DBReplaceLineFeedWithSpace(pAlias & "." & pRestrictions("address").Name)
      End If
      vWhere = " " & mvEnv.Connection.WhereClause(pRestrictions)
      If pAlias.Length > 0 AndAlso vWhere.Length > 0 Then
        For x As Integer = 1 To pRestrictions.Count
          Dim vRestrictionKey As String = pRestrictions.ItemKey(x)
          If vRestrictionKey <> "address" Then  'address field already has pAlias prefixed
            Dim vRestriction As CDBField = pRestrictions.Item(x)
            vWhere = Replace(vWhere, " " & vRestriction.Name & " ", " " & pAlias & "." & vRestriction.Name & " ")
          End If
        Next
      End If
      vWhere = Replace(vWhere, " contact_number ", " " & pAlias & ".contact_number ")
      vWhere = Replace(vWhere, " address_number ", " " & pAlias & ".address_number ")
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups And Len(mvEnv.GetConfig("ownership_list_manager")) > 0 Then
        'The use of the 'ownership_list_manager' config is taken from the MailingSelection.StepOwnershipRestriction method.
        'If this config isn't set then SELECT doesn't join to the 'ownership_group_users' table.
        vWhere = Replace(vWhere, " ownership_group ", " " & pAlias & ".ownership_group ")
      Else
        vWhere = Replace(vWhere, " department ", " " & pAlias & ".department ")
      End If
      vWhere = Replace(vWhere, " valid_from ", " " & pAlias & ".valid_from ")
      vWhere = Replace(vWhere, " valid_to ", " " & pAlias & ".valid_to ")
      vWhere = Replace(vWhere, " amended_on ", " " & pAlias & ".amended_on ")
      vWhere = Replace(vWhere, " amended_by ", " " & pAlias & ".amended_by ")

      Return vWhere
    End Function

    Public Sub LMCreateListTable()
      Dim vFields As New CDBFields

      LMDropListTable()
      With vFields
        .Add("selection_set", CDBField.FieldTypes.cftLong)
        .Add("revision", CDBField.FieldTypes.cftInteger)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("marker", CDBField.FieldTypes.cftInteger)          'used for generating random data sample
      End With
      mvEnv.Connection.CreateTableFromFields(ListTableName, vFields)
    End Sub

    Public Sub LMCreateEbuListTable()
      Dim vFields As New CDBFields
      LMDropListTable()
      With vFields
        .Add("selection_set", CDBField.FieldTypes.cftLong)
        .Add("revision", CDBField.FieldTypes.cftInteger)
        .Add("exam_booking_unit_id", CDBField.FieldTypes.cftLong)
      End With
      mvEnv.Connection.CreateTableFromFields(ListTableName, vFields)
    End Sub

    Public Sub LMDropListTable()

      If mvEnv.Connection.TableExists(ListTableName) Then
        mvEnv.Connection.DropTable(ListTableName)
      End If
    End Sub

    Public Sub ProcessBulkEmail(ByVal pTableName As String, ByRef pBulkEmailTableName As String, ByRef pNonEmailTableName As String, ByRef pEmailCount As Integer, ByRef pUsageCode As String)
      Dim vWhereFields As CDBFields
      Dim vSQL As String
      Dim vTempTable As String
      Dim vTempTable2 As String
      'First create the new table which will hold the communications number

      'Build the bulk email table
      vWhereFields = New CDBFields
      With vWhereFields
        .Add("selection_set", CDBField.FieldTypes.cftLong)
        .Add("revision", CDBField.FieldTypes.cftInteger)
        If mvMasterAttribute <> mvContactAttribute Then .Add(mvMasterAttribute, CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
        .Add("communication_number", CDBField.FieldTypes.cftLong)
      End With
      pBulkEmailTableName = pTableName & "_bem"
      mvConn.CreateTableFromFields(pBulkEmailTableName, vWhereFields)

      vTempTable = pTableName & "_be1"
      vWhereFields.Add("sequence", CDBField.FieldTypes.cftLong)
      mvConn.CreateTableFromFields(vTempTable, vWhereFields)

      'Select the email addresses by devious and diverse means
      If Len(pUsageCode) > 0 Then AddBulkEmailRecords(pTableName, vTempTable, BulkEmailSelectionTypes.bestByUsage, pUsageCode)
      AddBulkEmailRecords(pTableName, vTempTable, BulkEmailSelectionTypes.bestByPreferred, "")
      AddBulkEmailRecords(pTableName, vTempTable, BulkEmailSelectionTypes.bestByDefault, "")
      AddBulkEmailRecords(pTableName, vTempTable, BulkEmailSelectionTypes.bestByAny, "")

      vTempTable2 = pTableName & "_be2"
      vWhereFields.Remove("communication_number")
      mvConn.CreateTableFromFields(vTempTable2, vWhereFields)

      'Now select the lowest sequence number into temp table 2
      vSQL = "INSERT INTO " & vTempTable2 & " (selection_set, revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
      vSQL = vSQL & "contact_number, address_number, sequence) SELECT selection_set, revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
      vSQL = vSQL & "contact_number, address_number ,MIN(sequence) "
      vSQL = vSQL & "FROM " & vTempTable
      vSQL = vSQL & " GROUP BY selection_set, revision ,contact_number, address_number "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & "," & mvMasterAttribute
      mvConn.LogMailSQL(vSQL)
      mvConn.ExecuteSQL(vSQL)

      'Now get the email addresses
      vSQL = "INSERT INTO " & pBulkEmailTableName & " (selection_set, revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
      vSQL = vSQL & "contact_number, address_number, communication_number) SELECT st.selection_set, st.revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & "st." & mvMasterAttribute & ", "
      vSQL = vSQL & "st.contact_number, st.address_number ,st.communication_number "
      vSQL = vSQL & "FROM " & vTempTable & " st, " & vTempTable2 & " st2 WHERE "
      vSQL = vSQL & "st.selection_set = st2.selection_set AND st.revision = st2.revision "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & " AND st." & mvMasterAttribute & " = st2." & mvMasterAttribute
      vSQL = vSQL & " AND st.contact_number = st2.contact_number AND st.address_number = st2.address_number AND st.sequence = st2.sequence"
      mvConn.LogMailSQL(vSQL)
      pEmailCount = mvConn.ExecuteSQL(vSQL)

      mvConn.DropTable(vTempTable)
      mvConn.DropTable(vTempTable2)

      mvConn.CreateIndex(False, pBulkEmailTableName, {"selection_set", "revision", "contact_number"})

      'Build the non bulk email table
      vWhereFields = New CDBFields
      With vWhereFields
        .Add("selection_set", CDBField.FieldTypes.cftLong)
        .Add("revision", CDBField.FieldTypes.cftInteger)
        If mvMasterAttribute <> mvContactAttribute Then .Add(mvMasterAttribute, CDBField.FieldTypes.cftLong)
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("address_number", CDBField.FieldTypes.cftLong)
      End With
      pNonEmailTableName = pTableName & "_nem"
      mvConn.CreateTableFromFields(pNonEmailTableName, vWhereFields)

      'Now insert into the non email table all the contacts that are not in the email table
      vSQL = "INSERT INTO " & pNonEmailTableName & " (selection_set, revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
      vSQL = vSQL & "contact_number, address_number) SELECT st.selection_set, st.revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & "st." & mvMasterAttribute & ", "
      vSQL = vSQL & "st.contact_number, st.address_number "
      vSQL = vSQL & "FROM " & pTableName & " st LEFT OUTER JOIN " & pBulkEmailTableName & " et ON "
      vSQL = vSQL & "st.selection_set = et.selection_set AND st.revision = et.revision AND st.contact_number = et.contact_number "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & "AND st." & mvMasterAttribute & " = et." & mvMasterAttribute
      vSQL = vSQL & " WHERE et.communication_number IS NULL"
      vSQL = mvConn.ProcessAnsiJoins(vSQL)
      mvConn.LogMailSQL(vSQL)
      mvConn.ExecuteSQL(vSQL)
    End Sub

    Private Sub AddBulkEmailRecords(ByRef pSourceTableName As String, ByRef pDestTableName As String, ByRef pType As BulkEmailSelectionTypes, ByRef pUsage As String)
      Dim vSQL As String

      'Now insert into the email table all the contacts that have valid email addresses
      vSQL = "INSERT INTO " & pDestTableName & " (selection_set, revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
      vSQL = vSQL & "contact_number, address_number, communication_number, sequence) SELECT selection_set, revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
      vSQL = vSQL & "st.contact_number, st.address_number, MIN(co.communication_number) AS communication_number, " & pType
      vSQL = vSQL & " FROM " & pSourceTableName & " st, communications co, devices d "
      If pType = BulkEmailSelectionTypes.bestByUsage Then vSQL = vSQL & ", contact_communication_usages ccu "
      vSQL = vSQL & "WHERE st.contact_number = co.contact_number"
      vSQL = vSQL & " AND co.device = d.device AND d.email = 'Y' AND co.mail = 'Y' AND co.is_active = 'Y' "
      Select Case pType
        Case BulkEmailSelectionTypes.bestByUsage
          vSQL = vSQL & " AND co.communication_number = ccu.communication_number AND ccu.communication_usage = '" & pUsage & "' "
        Case BulkEmailSelectionTypes.bestByPreferred
          vSQL = vSQL & " AND co.preferred_method = 'Y'"
        Case BulkEmailSelectionTypes.bestByDefault
          vSQL = vSQL & " AND co.preferred_method = 'N' AND co.device_default = 'Y'"
        Case BulkEmailSelectionTypes.bestByAny
          vSQL = vSQL & " AND co.preferred_method = 'N' AND co.device_default = 'N'"
      End Select
      vSQL = vSQL & "GROUP BY selection_set, revision, "
      If mvMasterAttribute <> mvContactAttribute Then vSQL = vSQL & mvMasterAttribute & ", "
      vSQL = vSQL & "st.contact_number, st.address_number"
      mvConn.LogMailSQL(vSQL)
      mvConn.ExecuteSQL(vSQL)
    End Sub

    Public Sub CreateIrishGiftAidCertificates(ByVal pSelectionSet As Integer, ByVal pRevision As Integer)
      Dim vParams As New CDBParameters
      Dim vCertificate As GaAppropriateCertificate
      Dim vRS As CDBRecordSet
      Dim vAdd As Boolean
      Dim vEndDate As String
      Dim vStartDate As String
      Dim vSQL As String
      Dim vTableName As String

      'Selected data will include ContactNumber,AddressNumber,Performance
      vTableName = MailingTableName("Smcam", "Smapp")

      'Set the previous tax year
      vStartDate = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGATaxYearStart)
      If Len(vStartDate) = 0 Then vStartDate = "01/01/2008" 'just in case
      vStartDate = CStr(DateSerial(Year(CDate(TodaysDate())), Month(CDate(vStartDate)), Day(CDate(vStartDate))))
      If CDate(vStartDate) > CDate(TodaysDate()) Then vStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, CDate(vStartDate))) 'StartDate must be before Today
      vEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vStartDate))))

      If CDate(vEndDate) >= CDate(TodaysDate()) Then
        'This is the current tax year
        vEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, CDate(vEndDate)))
        vStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, -1, DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vEndDate))))
      End If

      'We do not know the TaxStatus of the donors as it is only stored on the AppropriateCertificate
      'so we will attempt to find a previous AppropriateCertificate and take that TaxStatus,
      'otherwise default to 'S' (Standard)
      vSQL = "SELECT ms.contact_number, cp.value_of_payments," & mvEnv.Connection.DBIsNull("ac.tax_status", "'S'") & " AS tax_status, ac.start_date, ac.cancellation_reason" ', {fn ifnull(ap.certificate_count,0)} AS certificate_count"
      vSQL = vSQL & " FROM " & vTableName & " ms"
      vSQL = vSQL & " INNER JOIN contact_performances cp ON ms.contact_number = cp.contact_number AND ms.performance = cp.performance"
      vSQL = vSQL & " INNER JOIN contacts c ON cp.contact_number = c.contact_number"
      vSQL = vSQL & " LEFT OUTER JOIN (SELECT gc.contact_number, gc.tax_status, cancellation_reason, claim_number, MAX(start_date) AS start_date FROM " & vTableName & " tt"
      vSQL = vSQL & " INNER JOIN ga_appropriate_certificates gc ON tt.contact_number = gc.contact_number" 'WHERE gc.cancellation_reason IS NULL"
      vSQL = vSQL & " GROUP BY gc.contact_number, gc.tax_status, cancellation_reason, claim_number) ac ON c.contact_number = ac.contact_number"
      vSQL = vSQL & " WHERE ms.selection_set = " & pSelectionSet & " AND revision = " & pRevision
      vSQL = vSQL & " AND cp.value_of_payments >= " & Val(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMinimumAnnualDonation))
      vSQL = vSQL & " AND c.contact_type = 'C' AND c.ni_number IS NOT NULL"

      vRS = mvEnv.Connection.GetRecordSetAnsiJoins(vSQL)

      If vRS.Fetch() = True Then
        With vParams
          .Add("ContactNumber", CDBField.FieldTypes.cftLong)
          .Add("CertificateAmount", CDBField.FieldTypes.cftNumeric)
          .Add("TaxStatus")
          .Add("StartDate", CDBField.FieldTypes.cftDate, vStartDate)
          .Add("EndDate", CDBField.FieldTypes.cftDate, vEndDate)
        End With
        Do
          vAdd = False
          If IsDate(vRS.Fields(4).Value) Then 'StartDate
            If (DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vRS.Fields(4).Value))) < CDate(vStartDate)) Then vAdd = True
            'If vAdd = True And Len(vRS.Fields(5).Value) > 0 Then vAdd = False         'CancellationReason
          Else
            vAdd = True
          End If
          If vAdd Then
            vParams(1).Value = vRS.Fields(1).IntegerValue.ToString 'ContactNumber
            vParams(2).Value = vRS.Fields(2).DoubleValue.ToString 'CertificateAmount
            vParams(3).Value = vRS.Fields(3).Value 'TaxStatus
            vCertificate = New GaAppropriateCertificate(mvEnv)
            With vCertificate
              vCertificate.Init()
              .Create(vParams)
              .Save()
            End With
          End If
        Loop While vRS.Fetch() = True
      End If
      vRS.CloseRecordSet()

    End Sub

    Public Sub CancelSOs(ByVal pSelectionTable As String, ByVal pCancReason As String, ByVal pSortOrder As MailingSelection.SortOrderTypes, ByVal pSelectionSet As String)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vPaymentPlan As PaymentPlan
      Dim vFields As String
      Dim vCancReasonDesc As String

      vPaymentPlan = New PaymentPlan
      vPaymentPlan.Init(mvEnv)
      vFields = vPaymentPlan.GetRecordSetFields(PaymentPlan.PayPlanRecordSetTypes.pprstNumbers Or PaymentPlan.PayPlanRecordSetTypes.pprstType Or PaymentPlan.PayPlanRecordSetTypes.pprstPPDetails)
      vCancReasonDesc = mvEnv.GetDescription("cancellation_reasons", "cancellation_reason", pCancReason)

      vSQL = "SELECT " & vFields
      vSQL = vSQL & " FROM " & mvSelectionTable & " sc, addresses a, countries co, contacts c, bankers_orders bo, bank_accounts ba, orders o, payment_frequencies pf, contact_accounts ca, banks b"
      vSQL = vSQL & " WHERE sc.selection_set = " & pSelectionSet
      vSQL = vSQL & " AND sc.revision = 1"
      vSQL = vSQL & " AND sc.address_number = a.address_number"
      vSQL = vSQL & " AND a.country = co.country"
      vSQL = vSQL & " AND sc.contact_number = c.contact_number"
      vSQL = vSQL & " AND sc.bankers_order_number = bo.bankers_order_number"
      vSQL = vSQL & " AND bo.cancellation_reason IS NULL"
      vSQL = vSQL & " AND bo.bank_account = ba.bank_account"
      vSQL = vSQL & " AND bo.order_number = o.order_number"
      vSQL = vSQL & " AND o.payment_frequency = pf.payment_frequency"
      vSQL = vSQL & " AND bo.bank_details_number = ca.bank_details_number"
      vSQL = vSQL & " AND ca.sort_code = b.sort_code"

      vRecordSet = mvConn.GetRecordSet(vSQL)
      With vRecordSet
        While .Fetch() = True
          vPaymentPlan = New PaymentPlan
          vPaymentPlan.InitFromRecordSet(mvEnv, vRecordSet, PaymentPlan.PayPlanRecordSetTypes.pprstNumbers Or PaymentPlan.PayPlanRecordSetTypes.pprstType Or PaymentPlan.PayPlanRecordSetTypes.pprstPPDetails)
          vPaymentPlan.Cancel(PaymentPlan.PaymentPlanCancellationTypes.pctStandingOrder Or PaymentPlan.PaymentPlanCancellationTypes.pctPaymentPlan, pCancReason, "", vCancReasonDesc, mvEnv.User.Logname, "")
        End While
        .CloseRecordSet()
      End With
      'GenerateSummaryPrint(pSortOrder)
    End Sub

    Public Sub CancelGayePledges(ByVal pSelectionTable As String, ByVal pCancReason As String, ByVal pSortOrder As MailingSelection.SortOrderTypes, ByVal pSelectionSet As String)
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String
      Dim vGAYEPledge As PreTaxPledge
      Dim vFields As String

      vGAYEPledge = New PreTaxPledge(mvEnv)
      vGAYEPledge.Init()
      vFields = vGAYEPledge.GetRecordSetFields()

      vSQL = "SELECT " & vFields
      vSQL = vSQL & " FROM " & pSelectionTable & " sc, addresses a, countries co, contacts c, gaye_pledges gp"
      vSQL = vSQL & " WHERE sc.selection_set = " & pSelectionSet
      vSQL = vSQL & " AND sc.revision = 1"
      vSQL = vSQL & " AND sc.address_number = a.address_number"
      vSQL = vSQL & " AND a.country = co.country"
      vSQL = vSQL & " AND sc.contact_number = c.contact_number"
      vSQL = vSQL & " AND sc.gaye_pledge_number = gp.gaye_pledge_number"

      vRecordSet = mvConn.GetRecordSet(vSQL)
      With vRecordSet
        While .Fetch() = True
          vGAYEPledge = New PreTaxPledge(mvEnv)
          vGAYEPledge.InitFromRecordSet(vRecordSet)
          vGAYEPledge.Cancel(pCancReason, "")
        End While
        .CloseRecordSet()
      End With
      'GenerateSummaryPrint(pSortOrder)
    End Sub


    Private Sub GenerateSummaryPrint(ByRef pSortOrder As MailingSelection.SortOrderTypes)
      'Dim vErrorNumber As Integer
      'Dim vSortFields As String

      'mvJob.InfoMessage = XLAT("Producing Contact Summary")
      ''Now print report
      'vSortFields = DetermineSortOrder(pSortOrder, False, False)
      'GMPrintSummary(vSortFields, "")
      'mvJob.InfoMessage = Caption & " - Summary Print"
    End Sub

    Sub GMPrintSummary(ByVal pSortFields As String, ByVal pTitle As String)
      'Dim vParam As CDBParameter
      'Dim vParams As New CDBParameters
      'Dim vType As String
      'Dim vDestType As CDBReports.Report.outDestTypes

      'vParams.Add("1", CDBField.FieldTypes.cftCharacter, SelectionTable)
      'vParams.Add("2", mvSelectionSet)
      'vParams.Add("3", 1) 'Revision
      'If mvJob.TaskJobType <> JobSchedule.TaskJobTypes.tjtPayrollPledgeCancellation And mvJob.TaskJobType <> JobSchedule.TaskJobTypes.tjtStandingOrderCancellation Then
      '  vParams.Add("4", CDBField.FieldTypes.cftCharacter, pSortFields)
      '  If Len(pTitle) = 0 Then
      '    Select Case mvJob.TaskJobType
      '      Case JobSchedule.TaskJobTypes.tjtMemberMailing, JobSchedule.TaskJobTypes.tjtMembCardMailing
      '      Case Else
      '        pTitle = "Summary Print"
      '    End Select
      '  End If
      '  vParams.Add("5", CDBField.FieldTypes.cftCharacter, pTitle)
      'Else
      '  vType = "CN"
      'End If

      'If MailingType = MailingTypes.mtyMembershipCards Then
      '  vType = "MC" & vType
      'Else
      '  vType = MailingTypeCode & vType
      'End If
      'vDestType = TaskUtils.DestinationFromDescription((mvJob.Parameters("ReportDestination").Value))
      'If vDestType <> CDBReports.Report.outDestTypes.odtNone Then
      '  mvReport = New CDBReports.Report
      '  mvReport.Load(mvEnv, 0, vType, mvEnv.ClientCode)
      '  For Each vParam In vParams
      '    mvReport.SetParameter((vParam.Name), (vParam.Value))
      '  Next vParam
      '  mvReport.Run(mvConn, CDBReports.Report.outDestTypes.odtSave, mvReportDest)
      'End If
    End Sub


    Public Sub LoadVariableCriteria(ByVal vParams As CDBParameters)
      Dim vListOfVariables As ArrayList
      Dim vParam As CDBParameter
      Dim vVariables As CDBParameters = Nothing
      Dim vName As String
      Dim vRange As Boolean
      Dim vVariableParam As CDBParameter
      Dim vVariableCriteria As VariableCriteria

      vListOfVariables = New ArrayList(vParams.Count)
      'Move any parameters that are actually criteria variables into a separate CDBParameters collection
      For Each vParam In vParams
        If Left(vParam.Name, 2) = "C_" Then
          If vVariables Is Nothing Then vVariables = New CDBParameters
          vName = "$" & Mid(vParam.Name, 3)
          vName = Replace(Replace(vName, "_M", "-"), "_P", "+")
          If IsNumeric(Right(vName, 1)) Then vRange = vVariables.Exists(Mid(vName, 1, Len(vName) - 1)) 'if the right-most character of the parameter's name is numeric then we may have encountered a range of values
          If vRange Then
            vName = Mid(vName, 1, Len(vName) - 1)
            vVariables.Item(vName).let_Value(vVariables.Item(vName).Value & " to " & vParam.Value)
          Else
            vVariables.Add(vName, vParam.DataType, vParam.Value)
          End If
          'vParams.Remove((vParam.Name))
          vListOfVariables.Add(vParam.Name)
        End If
      Next vParam

      ' Remove all variables from the parameterlist
      For vIndex As Integer = 0 To vListOfVariables.Count - 1
        vParams.Remove(vListOfVariables(vIndex).ToString)
      Next

      If vVariables IsNot Nothing AndAlso vVariables.Count > 0 Then
        InitVariableCriteria((CriteriaSetNumber))
        For Each vVariableParam In vVariables
          For Each vVariableCriteria In mvVariableCriteria
            If vVariableCriteria.VariableName = vVariableParam.Name And vVariableCriteria.Valid Then
              vVariableCriteria.Value = vVariableParam.Value
              Exit For
            End If
          Next vVariableCriteria
        Next vVariableParam
      End If
    End Sub

  End Class
End Namespace

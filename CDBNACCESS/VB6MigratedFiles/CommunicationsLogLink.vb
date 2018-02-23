

Namespace Access
  Public Class CommunicationsLogLink

    Public Enum CommunicationsLogLinkRecordSetTypes 'These are bit values
      cllrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CommunicationsLogLinkFields
      cllfAll = 0
      cllfCommunicationsLogNumber
      cllfContactNumber
      cllfAddressNumber
      cllfLinkType
      cllfProcessed
      cllfNotified
    End Enum

    Private Enum CommunicationsLogDocLinkFields
      cldlfAll = 0
      cldlfDocumentNumber1
      cldlfDocumentNumber2
      cldlfAmendedOn
      cldlfAmendedBy
    End Enum

    Private Enum CommunicationsLogTransLinkFields
      cltlfAll = 0
      cltlfDocumentNumber
      cltlfBatchNumber
      cltlfTransactionNumber
    End Enum

    Private Enum CommunicationsLogEventLinkFields
      clelfAll = 0
      clelfDocumentNumber
      clelfEventNumber
      clelfAmendedOn
      clelfAmendedBy
    End Enum

    Public Enum DocumentLinkTypes
      dltDocumentToContact
      dltDocumentToDocument
      dltDocumentToTransaction
      dltDocumentToEvent
      dltDocumentToExamCentre
      dltDocumentToExamUnit
      dltDocumentToExamCentreUnit
      dltDocumentToFundraisingRequest
      dltDocumentToCPDPeriod
      dltDocumentToCPDPoint
      dltDocumentToContactPosition
    End Enum

    Private Enum CommsLogExamCentreLinkFields
      clelfAll = 0
      clelfDocumentNumber
      clelfExamCentreId
      clelfLinkType
      clelfDocumentLinkId
      clelfAmendedOn
      clelfAmendedBy
    End Enum

    Private Enum CommsLogFundraisingRequestLinkFields
      All = 0
      clfrlDocumentNumber
      clfrlFundraisingRequestNumber
      clfrlLinkType
      clfrlDocumentLinkId
      clfrlAmendedOn
      clfrlAmendedBy
    End Enum

    Private Enum CommsLogCPDLinkFields
      All = 0
      DocumentNumber
      CPDPeriodOrPointNumber
      LinkType
      DocumentLinkId
      AmendedOn
      AmendedBy
    End Enum
    Private Enum CommsLogPositionLinkFields
      All = 0
      DocumentNumber
      ContactPositionNumber
      LinkType
      DocumentLinkId
      AmendedOn
      AmendedBy
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean
    Private mvDocumentLink As Boolean
    Private mvTransactionLink As Boolean
    Private mvEventLink As Boolean
    Private mvExamCentreId As Boolean
    Private mvExamUnitId As Boolean
    Private mvExamCentreUnitId As Boolean
    Private mvFundraisingRequestNo As Boolean
    Private mvCPDPeriodLink As Boolean
    Private mvCPDPointLink As Boolean
    Private mvPositionLink As Boolean

    Public Enum CommunicationLogLinkTypes
      clltNone
      clltAddressee
      clltSender
      clltCopied
      clltDistributed
      clltRelated
    End Enum
    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          If mvDocumentLink Then
            .DatabaseTableName = "communications_log_doc_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number_1", CDBField.FieldTypes.cftLong)
            .Add("communications_log_number_2", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommunicationsLogDocLinkFields.cldlfDocumentNumber1).SetPrimaryKeyOnly()
            .Item(CommunicationsLogDocLinkFields.cldlfDocumentNumber2).SetPrimaryKeyOnly()
          ElseIf mvEventLink Then
            .DatabaseTableName = "event_documents"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("event_number", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommunicationsLogEventLinkFields.clelfDocumentNumber).SetPrimaryKeyOnly()
            .Item(CommunicationsLogEventLinkFields.clelfEventNumber).SetPrimaryKeyOnly()
          ElseIf mvTransactionLink Then
            .DatabaseTableName = "communications_log_trans"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("batch_number", CDBField.FieldTypes.cftLong)
            .Add("transaction_number", CDBField.FieldTypes.cftLong)

            .Item(CommunicationsLogTransLinkFields.cltlfDocumentNumber).SetPrimaryKeyOnly()
            .Item(CommunicationsLogTransLinkFields.cltlfBatchNumber).SetPrimaryKeyOnly()
            .Item(CommunicationsLogTransLinkFields.cltlfTransactionNumber).SetPrimaryKeyOnly()
          ElseIf mvExamCentreId Then
            .DatabaseTableName = "document_log_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("exam_centre_id", CDBField.FieldTypes.cftLong)
            .Add("link_type")
            .Add("document_link_id", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommsLogExamCentreLinkFields.clelfDocumentNumber).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfExamCentreId).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfLinkType).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfDocumentLinkId).SetPrimaryKeyOnly()
          ElseIf mvExamUnitId Then
            .DatabaseTableName = "document_log_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("exam_unit_link_id", CDBField.FieldTypes.cftLong)
            .Add("link_type")
            .Add("document_link_id", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommsLogExamCentreLinkFields.clelfDocumentNumber).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfExamCentreId).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfLinkType).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfDocumentLinkId).SetPrimaryKeyOnly()
          ElseIf mvExamCentreUnitId Then
            .DatabaseTableName = "document_log_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("exam_centre_unit_id", CDBField.FieldTypes.cftLong)
            .Add("link_type")
            .Add("document_link_id", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommsLogExamCentreLinkFields.clelfDocumentNumber).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfExamCentreId).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfLinkType).SetPrimaryKeyOnly()
            .Item(CommsLogExamCentreLinkFields.clelfDocumentLinkId).SetPrimaryKeyOnly()
          ElseIf mvFundraisingRequestNo Then
            .DatabaseTableName = "document_log_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("fundraising_request_number")
            .Add("link_type")
            .Add("document_link_id", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommsLogFundraisingRequestLinkFields.clfrlDocumentNumber).SetPrimaryKeyOnly()
            .Item(CommsLogFundraisingRequestLinkFields.clfrlFundraisingRequestNumber).SetPrimaryKeyOnly()
            .Item(CommsLogFundraisingRequestLinkFields.clfrlLinkType).SetPrimaryKeyOnly()
            .Item(CommsLogFundraisingRequestLinkFields.clfrlDocumentLinkId).SetPrimaryKeyOnly()
          ElseIf mvCPDPeriodLink = True OrElse mvCPDPointLink = True Then
            .DatabaseTableName = "document_log_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            If mvCPDPeriodLink Then
              .Add("contact_cpd_period_number", CDBField.FieldTypes.cftLong)
            ElseIf mvCPDPointLink Then
              .Add("contact_cpd_point_number", CDBField.FieldTypes.cftLong)
            End If
            .Add("link_type")
            .Add("document_link_id", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommsLogCPDLinkFields.DocumentNumber).SetPrimaryKeyOnly()
            .Item(CommsLogCPDLinkFields.CPDPeriodOrPointNumber).SetPrimaryKeyOnly()
            .Item(CommsLogCPDLinkFields.LinkType).SetPrimaryKeyOnly()
            .Item(CommsLogCPDLinkFields.DocumentLinkId).SetPrimaryKeyOnly()
          ElseIf mvPositionLink = True Then
            .DatabaseTableName = "document_log_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("contact_position_number", CDBField.FieldTypes.cftLong)
            .Add("link_type")
            .Add("document_link_id", CDBField.FieldTypes.cftLong)
            .Add("amended_on", CDBField.FieldTypes.cftDate)
            .Add("amended_by")

            .Item(CommsLogPositionLinkFields.DocumentNumber).SetPrimaryKeyOnly()
            .Item(CommsLogPositionLinkFields.ContactPositionNumber).SetPrimaryKeyOnly()
            .Item(CommsLogPositionLinkFields.LinkType).SetPrimaryKeyOnly()
            .Item(CommsLogPositionLinkFields.DocumentLinkId).SetPrimaryKeyOnly()
          Else
            .DatabaseTableName = "communications_log_links"
            'There should be an entry here for each field in the table
            'Keep these in the same order as the Fields enum
            .Add("communications_log_number", CDBField.FieldTypes.cftLong)
            .Add("contact_number", CDBField.FieldTypes.cftLong)
            .Add("address_number", CDBField.FieldTypes.cftLong)
            .Add("link_type")
            .Add("processed")
            .Add("notified")

            .Item(CommunicationsLogLinkFields.cllfCommunicationsLogNumber).SetPrimaryKeyOnly()
            .Item(CommunicationsLogLinkFields.cllfContactNumber).SetPrimaryKeyOnly()
            .Item(CommunicationsLogLinkFields.cllfLinkType).SetPrimaryKeyOnly()
          End If
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CommunicationsLogLinkFields)
      'Add code here to ensure all values are valid before saving
      Dim vAmendedByIndex As Integer = 0
      Dim vAmendedOnIndex As Integer = 0
      If mvDocumentLink Or mvEventLink Then
        vAmendedByIndex = CommunicationsLogDocLinkFields.cldlfAmendedBy
        vAmendedOnIndex = CommunicationsLogDocLinkFields.cldlfAmendedOn
      ElseIf mvExamCentreId Or mvExamCentreUnitId Or mvExamUnitId Or mvFundraisingRequestNo Then
        vAmendedByIndex = CommsLogExamCentreLinkFields.clelfAmendedBy
        vAmendedOnIndex = CommsLogExamCentreLinkFields.clelfAmendedOn
      ElseIf mvCPDPeriodLink OrElse mvCPDPointLink Then
        vAmendedByIndex = CommsLogCPDLinkFields.AmendedBy
        vAmendedOnIndex = CommsLogCPDLinkFields.AmendedOn
      ElseIf mvPositionLink Then
        vAmendedByIndex = CommsLogPositionLinkFields.AmendedBy
        vAmendedOnIndex = CommsLogPositionLinkFields.AmendedOn
      End If
      If vAmendedByIndex > 0 Then
        mvClassFields(vAmendedByIndex).Value = mvEnv.User.UserID
        mvClassFields(vAmendedOnIndex).Value = TodaysDate()
      End If
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CommunicationsLogLinkRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CommunicationsLogLinkRecordSetTypes.cllrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "cll")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByVal pCommunicationsLogNumber As Integer = 0, Optional ByVal pLinkType As CommunicationsLogLink.CommunicationLogLinkTypes = CommunicationsLogLink.CommunicationLogLinkTypes.clltNone, Optional ByVal pContactDocumentBatchOrEventNumber As Integer = 0, Optional ByVal pDocumentLinkType As CommunicationsLogLink.DocumentLinkTypes = CommunicationsLogLink.DocumentLinkTypes.dltDocumentToContact, Optional ByVal pTransactionNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      Select Case pDocumentLinkType
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToDocument
          mvDocumentLink = True
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToTransaction
          mvTransactionLink = True
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToEvent
          mvEventLink = True
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToExamCentre
          mvExamCentreId = True
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToExamUnit
          mvExamUnitId = True
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToExamCentreUnit
          mvExamCentreUnitId = True
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToFundraisingRequest
          mvFundraisingRequestNo = True
        Case DocumentLinkTypes.dltDocumentToCPDPeriod
          mvCPDPeriodLink = True
        Case DocumentLinkTypes.dltDocumentToCPDPoint
          mvCPDPointLink = True
        Case DocumentLinkTypes.dltDocumentToContactPosition
          mvPositionLink = True
      End Select

      InitClassFields()
      If pCommunicationsLogNumber > 0 Then
        vWhereFields.Add((mvClassFields(CommunicationsLogLinkFields.cllfCommunicationsLogNumber).Name), CDBField.FieldTypes.cftLong, pCommunicationsLogNumber)
        If pLinkType <> CommunicationsLogLink.CommunicationLogLinkTypes.clltNone Then vWhereFields.Add("link_type", CDBField.FieldTypes.cftCharacter, SetLinkType(pLinkType))
        If pContactDocumentBatchOrEventNumber > 0 Then vWhereFields.Add((mvClassFields(CommunicationsLogLinkFields.cllfContactNumber).Name), CDBField.FieldTypes.cftLong, pContactDocumentBatchOrEventNumber)
        If pTransactionNumber > 0 Then vWhereFields.Add((mvClassFields(CommunicationsLogTransLinkFields.cltlfTransactionNumber).Name), CDBField.FieldTypes.cftLong, pTransactionNumber)
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CommunicationsLogLinkRecordSetTypes.cllrtAll) & " FROM " & mvClassFields.DatabaseTableName & " cll WHERE " & pEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CommunicationsLogLinkRecordSetTypes.cllrtAll)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CommunicationsLogLinkRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        'Modify below to handle each recordset type as required
        If (pRSType And CommunicationsLogLinkRecordSetTypes.cllrtAll) = CommunicationsLogLinkRecordSetTypes.cllrtAll Then
          If mvDocumentLink Then
            .SetItem(CommunicationsLogDocLinkFields.cldlfDocumentNumber1, vFields)
            .SetItem(CommunicationsLogDocLinkFields.cldlfDocumentNumber2, vFields)
            .SetItem(CommunicationsLogDocLinkFields.cldlfAmendedOn, vFields)
            .SetItem(CommunicationsLogDocLinkFields.cldlfAmendedBy, vFields)
          ElseIf mvEventLink Then
            .SetItem(CommunicationsLogEventLinkFields.clelfDocumentNumber, vFields)
            .SetItem(CommunicationsLogEventLinkFields.clelfEventNumber, vFields)
            .SetItem(CommunicationsLogEventLinkFields.clelfAmendedOn, vFields)
            .SetItem(CommunicationsLogEventLinkFields.clelfAmendedBy, vFields)
          ElseIf mvTransactionLink Then
            .SetItem(CommunicationsLogTransLinkFields.cltlfDocumentNumber, vFields)
            .SetItem(CommunicationsLogTransLinkFields.cltlfBatchNumber, vFields)
            .SetItem(CommunicationsLogTransLinkFields.cltlfTransactionNumber, vFields)
          ElseIf mvExamCentreId Or mvExamCentreUnitId Or mvExamUnitId Then
            .SetItem(CommsLogExamCentreLinkFields.clelfDocumentNumber, vFields)
            .SetItem(CommsLogExamCentreLinkFields.clelfExamCentreId, vFields)
            .SetItem(CommsLogExamCentreLinkFields.clelfLinkType, vFields)
            .SetItem(CommsLogExamCentreLinkFields.clelfDocumentLinkId, vFields)
          Else
            .SetItem(CommunicationsLogLinkFields.cllfCommunicationsLogNumber, vFields)
            .SetItem(CommunicationsLogLinkFields.cllfContactNumber, vFields)
            .SetItem(CommunicationsLogLinkFields.cllfAddressNumber, vFields)
            .SetItem(CommunicationsLogLinkFields.cllfLinkType, vFields)
            .SetItem(CommunicationsLogLinkFields.cllfProcessed, vFields)
            .SetItem(CommunicationsLogLinkFields.cllfNotified, vFields)
          End If
        End If
      End With
    End Sub

    Public Sub DeleteAllLinks(ByVal pDocumentNumber As Integer)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add(mvClassFields(CommunicationsLogTransLinkFields.cltlfDocumentNumber).Name, CDBField.FieldTypes.cftInteger, pDocumentNumber)
      mvEnv.Connection.DeleteRecords(mvClassFields.DatabaseTableName, vWhereFields, False)
    End Sub

    Public Sub Delete(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
    End Sub

    Public Sub Save(Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      SetValid(CommunicationsLogLinkFields.cllfAll)
      If mvDocumentLink = False Then pAmendedBy = ""
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property BatchNumber() As Integer
      Get
        BatchNumber = mvClassFields.Item(CommunicationsLogTransLinkFields.cltlfBatchNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property TransactionNumber() As Integer
      Get
        TransactionNumber = mvClassFields.Item(CommunicationsLogTransLinkFields.cltlfTransactionNumber).IntegerValue
      End Get
    End Property

    Public Property AddressNumber() As Integer
      Get
        AddressNumber = mvClassFields.Item(CommunicationsLogLinkFields.cllfAddressNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CommunicationsLogLinkFields.cllfAddressNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property CommunicationsLogNumber() As Integer
      Get
        CommunicationsLogNumber = mvClassFields.Item(CommunicationsLogLinkFields.cllfCommunicationsLogNumber).IntegerValue
      End Get
    End Property

    Public Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(CommunicationsLogLinkFields.cllfContactNumber).IntegerValue
      End Get
      Set(ByVal Value As Integer)
        mvClassFields.Item(CommunicationsLogLinkFields.cllfContactNumber).IntegerValue = Value
      End Set
    End Property

    Public ReadOnly Property LinkType() As CommunicationsLogLink.CommunicationLogLinkTypes
      Get
        Select Case mvClassFields.Item(CommunicationsLogLinkFields.cllfLinkType).Value
          Case "A"
            LinkType = CommunicationsLogLink.CommunicationLogLinkTypes.clltAddressee
          Case "C"
            LinkType = CommunicationsLogLink.CommunicationLogLinkTypes.clltCopied
          Case "D"
            LinkType = CommunicationsLogLink.CommunicationLogLinkTypes.clltDistributed
          Case "R"
            LinkType = CommunicationsLogLink.CommunicationLogLinkTypes.clltRelated
          Case "S"
            LinkType = CommunicationsLogLink.CommunicationLogLinkTypes.clltSender
        End Select
      End Get
    End Property

    Public Shared Function GetLinkTypeCode(ByVal pLinkType As CommunicationsLogLink.CommunicationLogLinkTypes) As String
      Select Case pLinkType
        Case CommunicationLogLinkTypes.clltAddressee
          Return "A"
        Case CommunicationLogLinkTypes.clltCopied
          Return "C"
        Case CommunicationLogLinkTypes.clltDistributed
          Return "D"
        Case CommunicationLogLinkTypes.clltRelated
          Return "R"
        Case CommunicationLogLinkTypes.clltSender
          Return "S"
        Case Else
          Return "A"      'To fix compiler warning
      End Select
    End Function

    Public Property Notified() As Boolean
      Get
        Notified = mvClassFields.Item(CommunicationsLogLinkFields.cllfNotified).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(CommunicationsLogLinkFields.cllfNotified).Bool = Value
      End Set
    End Property

    Public Property Processed() As Boolean
      Get
        Processed = mvClassFields.Item(CommunicationsLogLinkFields.cllfProcessed).Bool
      End Get
      Set(ByVal Value As Boolean)
        mvClassFields.Item(CommunicationsLogLinkFields.cllfProcessed).Bool = Value
      End Set
    End Property

    Public Sub Create(ByRef pType As CommunicationsLogLink.DocumentLinkTypes, ByVal pEnv As CDBEnvironment, ByVal pCommsLogNumber As Integer, ByVal pContactDocumentBatchOrEventNumber As Integer, ByVal pAddressOrTransactionNumber As Integer, ByVal pLinkType As CommunicationsLogLink.CommunicationLogLinkTypes, ByVal pNotified As Boolean, ByVal pProcessed As Boolean)
      mvEnv = pEnv
      Select Case pType
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToDocument
          mvDocumentLink = True
        Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToTransaction
          mvTransactionLink = True
      End Select
      InitClassFields()
      With mvClassFields
        Select Case pType
          Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToDocument
            .Item(CommunicationsLogDocLinkFields.cldlfDocumentNumber1).Value = CStr(pCommsLogNumber)
            .Item(CommunicationsLogDocLinkFields.cldlfDocumentNumber2).Value = CStr(pContactDocumentBatchOrEventNumber)
          Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToEvent
            .Item(CommunicationsLogEventLinkFields.clelfDocumentNumber).Value = CStr(pCommsLogNumber)
            .Item(CommunicationsLogEventLinkFields.clelfEventNumber).Value = CStr(pContactDocumentBatchOrEventNumber)
          Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToTransaction
            .Item(CommunicationsLogTransLinkFields.cltlfDocumentNumber).Value = CStr(pCommsLogNumber)
            .Item(CommunicationsLogTransLinkFields.cltlfBatchNumber).Value = CStr(pContactDocumentBatchOrEventNumber)
            .Item(CommunicationsLogTransLinkFields.cltlfTransactionNumber).Value = CStr(pAddressOrTransactionNumber)
          Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToExamCentre, CommunicationsLogLink.DocumentLinkTypes.dltDocumentToExamCentreUnit, CommunicationsLogLink.DocumentLinkTypes.dltDocumentToExamUnit
            .Item(CommsLogExamCentreLinkFields.clelfDocumentNumber).Value = CStr(pCommsLogNumber)
            .Item(CommsLogExamCentreLinkFields.clelfExamCentreId).Value = CStr(pContactDocumentBatchOrEventNumber)
            .Item(CommsLogExamCentreLinkFields.clelfLinkType).Value = SetLinkType(pLinkType)
            .Item(CommsLogExamCentreLinkFields.clelfDocumentLinkId).Value = CStr(mvEnv.GetControlNumber("EDL"))
          Case CommunicationsLogLink.DocumentLinkTypes.dltDocumentToFundraisingRequest
            .Item(CommsLogFundraisingRequestLinkFields.clfrlDocumentNumber).Value = CStr(pCommsLogNumber)
            .Item(CommsLogFundraisingRequestLinkFields.clfrlFundraisingRequestNumber).Value = CStr(pContactDocumentBatchOrEventNumber)
            .Item(CommsLogFundraisingRequestLinkFields.clfrlLinkType).Value = SetLinkType(pLinkType)
            .Item(CommsLogFundraisingRequestLinkFields.clfrlDocumentLinkId).Value = CStr(mvEnv.GetControlNumber("EDL"))
          Case DocumentLinkTypes.dltDocumentToCPDPeriod, DocumentLinkTypes.dltDocumentToCPDPoint
            .Item(CommsLogCPDLinkFields.DocumentNumber).IntegerValue = pCommsLogNumber
            .Item(CommsLogCPDLinkFields.CPDPeriodOrPointNumber).IntegerValue = pContactDocumentBatchOrEventNumber
            .Item(CommsLogCPDLinkFields.LinkType).Value = SetLinkType(pLinkType)
            .Item(CommsLogCPDLinkFields.DocumentLinkId).IntegerValue = mvEnv.GetControlNumber("EDL")
          Case DocumentLinkTypes.dltDocumentToContactPosition
            .Item(CommsLogPositionLinkFields.DocumentNumber).IntegerValue = pCommsLogNumber
            .Item(CommsLogPositionLinkFields.ContactPositionNumber).IntegerValue = pContactDocumentBatchOrEventNumber
            .Item(CommsLogPositionLinkFields.LinkType).Value = SetLinkType(pLinkType)
            .Item(CommsLogPositionLinkFields.DocumentLinkId).IntegerValue = mvEnv.GetControlNumber("EDL")
          Case Else
            .Item(CommunicationsLogLinkFields.cllfCommunicationsLogNumber).Value = CStr(pCommsLogNumber)
            .Item(CommunicationsLogLinkFields.cllfContactNumber).Value = CStr(pContactDocumentBatchOrEventNumber)
            .Item(CommunicationsLogLinkFields.cllfAddressNumber).Value = CStr(pAddressOrTransactionNumber)
            .Item(CommunicationsLogLinkFields.cllfLinkType).Value = SetLinkType(pLinkType)
            .Item(CommunicationsLogLinkFields.cllfNotified).Bool = pNotified
            .Item(CommunicationsLogLinkFields.cllfProcessed).Bool = pProcessed
        End Select
      End With
      Save(mvEnv.User.UserID, True)
    End Sub

    Private Function SetLinkType(ByVal pLinkType As CommunicationsLogLink.CommunicationLogLinkTypes) As String
      Select Case pLinkType
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltAddressee
          Return "A"
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltCopied
          Return "C"
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltDistributed
          Return "D"
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltRelated
          Return "R"
        Case CommunicationsLogLink.CommunicationLogLinkTypes.clltSender
          Return "S"
        Case Else
          Return "A"    'To fix compiler warning
      End Select
    End Function
  End Class
End Namespace

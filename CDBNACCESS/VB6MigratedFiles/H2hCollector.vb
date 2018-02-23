

Namespace Access
  Public Class H2hCollector

    Public Enum H2hCollectorRecordSetTypes 'These are bit values
      hcortAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum H2hCollectorFields
      hcofAll = 0
      hcofCollectorNumber
      hcofCollectionNumber
      hcofContactNumber
      hcofRoute
      hcofRouteType
      hcofNoOfPremises
      hcofOperatorContactNumber
      hcofCollectorStatus
      hcofReadyForConfirmation
      hcofConfirmationProducedOn
      hcofReminderProducedOn
      hcofNotes
      hcofAmendedBy
      hcofAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "h2h_collectors"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collector_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("route")
          .Add("route_type")
          .Add("no_of_premises", CDBField.FieldTypes.cftLong)
          .Add("operator_contact_number", CDBField.FieldTypes.cftLong)
          .Add("collector_status")
          .Add("ready_for_confirmation")
          .Add("confirmation_produced_on", CDBField.FieldTypes.cftDate)
          .Add("reminder_produced_on", CDBField.FieldTypes.cftDate)
          .Add("notes")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(H2hCollectorFields.hcofCollectorNumber).SetPrimaryKeyOnly()
          .SetUniqueField(H2hCollectorFields.hcofCollectionNumber)
          .SetUniqueField(H2hCollectorFields.hcofContactNumber)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As H2hCollectorFields)
      'Add code here to ensure all values are valid before saving
      With mvClassFields
        If .Item(H2hCollectorFields.hcofCollectorNumber).IntegerValue = 0 Then .Item(H2hCollectorFields.hcofCollectorNumber).IntegerValue = mvEnv.GetControlNumber("CE")
        .Item(H2hCollectorFields.hcofAmendedOn).Value = TodaysDate()
        .Item(H2hCollectorFields.hcofAmendedBy).Value = mvEnv.User.Logname
      End With
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As H2hCollectorRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = H2hCollectorRecordSetTypes.hcortAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "hc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectorNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectorNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(H2hCollectorRecordSetTypes.hcortAll) & " FROM h2h_collectors hc WHERE collector_number = " & pCollectorNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, H2hCollectorRecordSetTypes.hcortAll)
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

    Public Sub InitFromCollectionAndContact(ByVal pEnv As CDBEnvironment, ByRef pCollectionNumber As Integer, ByRef pContactNumber As Integer)
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields

      mvEnv = pEnv
      With vWhereFields
        .Add("collection_number", CDBField.FieldTypes.cftCharacter, pCollectionNumber)
        .Add("contact_number", CDBField.FieldTypes.cftCharacter, pContactNumber)
      End With
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(H2hCollectorRecordSetTypes.hcortAll) & " FROM h2h_collectors hc WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, H2hCollectorRecordSetTypes.hcortAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As H2hCollectorRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(H2hCollectorFields.hcofCollectorNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And H2hCollectorRecordSetTypes.hcortAll) = H2hCollectorRecordSetTypes.hcortAll Then
          .SetItem(H2hCollectorFields.hcofCollectionNumber, vFields)
          .SetItem(H2hCollectorFields.hcofContactNumber, vFields)
          .SetItem(H2hCollectorFields.hcofRoute, vFields)
          .SetItem(H2hCollectorFields.hcofRouteType, vFields)
          .SetItem(H2hCollectorFields.hcofNoOfPremises, vFields)
          .SetItem(H2hCollectorFields.hcofOperatorContactNumber, vFields)
          .SetItem(H2hCollectorFields.hcofCollectorStatus, vFields)
          .SetItem(H2hCollectorFields.hcofReadyForConfirmation, vFields)
          .SetItem(H2hCollectorFields.hcofConfirmationProducedOn, vFields)
          .SetItem(H2hCollectorFields.hcofReminderProducedOn, vFields)
          .SetItem(H2hCollectorFields.hcofNotes, vFields)
          .SetItem(H2hCollectorFields.hcofAmendedBy, vFields)
          .SetItem(H2hCollectorFields.hcofAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(H2hCollectorFields.hcofAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(H2hCollectorFields.hcofCollectionNumber).Value = pParams("CollectionNumber").Value
        .Item(H2hCollectorFields.hcofContactNumber).Value = pParams("ContactNumber").Value
        .Item(H2hCollectorFields.hcofRoute).Value = pParams("Route").Value
        .Item(H2hCollectorFields.hcofRouteType).Value = pParams("RouteType").Value
        .Item(H2hCollectorFields.hcofNoOfPremises).Value = pParams("NoOfPremises").Value
        .Item(H2hCollectorFields.hcofOperatorContactNumber).Value = pParams("OperatorContactNumber").Value
        .Item(H2hCollectorFields.hcofCollectorStatus).Value = pParams("CollectorStatus").Value
        .Item(H2hCollectorFields.hcofReadyForConfirmation).Value = pParams("ReadyForConfirmation").Value
        .Item(H2hCollectorFields.hcofNotes).Value = pParams("Notes").Value
        If pParams.Exists("ConfirmationProducedOn") Then .Item(H2hCollectorFields.hcofConfirmationProducedOn).Value = pParams("ConfirmationProducedOn").Value
        If pParams.Exists("ReminderProducedOn") Then .Item(H2hCollectorFields.hcofReminderProducedOn).Value = pParams("ReminderProducedOn").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("Route") Then .Item(H2hCollectorFields.hcofRoute).Value = pParams("Route").Value
        If pParams.Exists("RouteType") Then .Item(H2hCollectorFields.hcofRouteType).Value = pParams("RouteType").Value
        If pParams.Exists("NoOfPremises") Then .Item(H2hCollectorFields.hcofNoOfPremises).Value = pParams("NoOfPremises").Value
        If pParams.Exists("OperatorContactNumber") Then .Item(H2hCollectorFields.hcofOperatorContactNumber).Value = pParams("OperatorContactNumber").Value
        If pParams.Exists("CollectorStatus") Then .Item(H2hCollectorFields.hcofCollectorStatus).Value = pParams("CollectorStatus").Value
        If pParams.Exists("ReadyForConfirmation") Then .Item(H2hCollectorFields.hcofReadyForConfirmation).Value = pParams("ReadyForConfirmation").Value
        If pParams.Exists("ConfirmationProducedOn") Then .Item(H2hCollectorFields.hcofConfirmationProducedOn).Value = pParams("ConfirmationProducedOn").Value
        If pParams.Exists("ReminderProducedOn") Then .Item(H2hCollectorFields.hcofReminderProducedOn).Value = pParams("ReminderProducedOn").Value
        If pParams.Exists("Notes") Then .Item(H2hCollectorFields.hcofNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vDS As New VBDataSelection
      Dim vDT As CDBDataTable
      Dim vParams As CDBParameters

      vParams = New CDBParameters
      vParams.Add("ContactNumber", ContactNumber)
      vParams.Add("CollectionNumber", CollectionNumber)
      vDS.Init(mvEnv, DataSelection.DataSelectionTypes.dstContactCollectionPayments, vParams)
      vDT = vDS.DataTable
      If vDT.Rows.Count() > 0 Then RaiseError(DataAccessErrors.daeCannotDeleteCollectorHasPayments)

      mvEnv.Connection.StartTransaction()
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
      mvEnv.Connection.ExecuteSQL("UPDATE collection_pis SET collector_number = NULL WHERE collector_number  = " & CollectorNumber)
      mvEnv.Connection.CommitTransaction()
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(H2hCollectorFields.hcofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(H2hCollectorFields.hcofAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        Return mvClassFields.Item(H2hCollectorFields.hcofCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectorNumber() As Integer
      Get
        Return mvClassFields.Item(H2hCollectorFields.hcofCollectorNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectorStatus() As String
      Get
        CollectorStatus = mvClassFields.Item(H2hCollectorFields.hcofCollectorStatus).Value
      End Get
    End Property

    Public ReadOnly Property ConfirmationProducedOn() As String
      Get
        ConfirmationProducedOn = mvClassFields.Item(H2hCollectorFields.hcofConfirmationProducedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields.Item(H2hCollectorFields.hcofContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property NoOfPremises() As Integer
      Get
        Return mvClassFields.Item(H2hCollectorFields.hcofNoOfPremises).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(H2hCollectorFields.hcofNotes).Value
      End Get
    End Property

    Public ReadOnly Property OperatorContactNumber() As Integer
      Get
        Return mvClassFields.Item(H2hCollectorFields.hcofOperatorContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ReadyForConfirmation() As String
      Get
        ReadyForConfirmation = mvClassFields.Item(H2hCollectorFields.hcofReadyForConfirmation).Value
      End Get
    End Property

    Public ReadOnly Property ReminderProducedOn() As String
      Get
        ReminderProducedOn = mvClassFields.Item(H2hCollectorFields.hcofReminderProducedOn).Value
      End Get
    End Property

    Public ReadOnly Property Route() As String
      Get
        Route = mvClassFields.Item(H2hCollectorFields.hcofRoute).Value
      End Get
    End Property

    Public ReadOnly Property RouteType() As String
      Get
        RouteType = mvClassFields.Item(H2hCollectorFields.hcofRouteType).Value
      End Get
    End Property
  End Class
End Namespace

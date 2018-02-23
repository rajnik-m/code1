

Namespace Access
  Public Class MannedCollector

    Public Enum MannedCollectorRecordSetTypes 'These are bit values
      mcortAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum MannedCollectorFields
      mcofAll = 0
      mcofCollectorNumber
      mcofCollectionNumber
      mcofContactNumber
      mcofTotalTime
      mcofAttended
      mcofReadyForConfirmation
      mcofReadyForAcknowledgement
      mcofConfirmationProducedOn
      mcofAcknowledgementProducedOn
      mcofNotes
      mcofAmendedBy
      mcofAmendedOn
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
          .DatabaseTableName = "manned_collectors"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collector_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("total_time")
          .Add("attended")
          .Add("ready_for_confirmation")
          .Add("ready_for_acknowledgement")
          .Add("confirmation_produced_on", CDBField.FieldTypes.cftDate)
          .Add("acknowledgement_produced_on", CDBField.FieldTypes.cftDate)
          .Add("notes")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)

          .Item(MannedCollectorFields.mcofCollectorNumber).SetPrimaryKeyOnly()

          .SetUniqueField(MannedCollectorFields.mcofCollectionNumber)
          .SetUniqueField(MannedCollectorFields.mcofContactNumber)
        End With
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As MannedCollectorFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(MannedCollectorFields.mcofCollectorNumber).IntegerValue = 0 Then mvClassFields.Item(MannedCollectorFields.mcofCollectorNumber).IntegerValue = mvEnv.GetControlNumber("CE")
      mvClassFields.Item(MannedCollectorFields.mcofAmendedOn).Value = TodaysDate()
      mvClassFields.Item(MannedCollectorFields.mcofAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As MannedCollectorRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = MannedCollectorRecordSetTypes.mcortAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "mc")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectorNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectorNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(MannedCollectorRecordSetTypes.mcortAll) & " FROM manned_collectors mc WHERE collector_number = " & pCollectorNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, MannedCollectorRecordSetTypes.mcortAll)
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
        .Add("collection_number", CDBField.FieldTypes.cftLong, pCollectionNumber)
        .Add("contact_number", CDBField.FieldTypes.cftLong, pContactNumber)
      End With
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(MannedCollectorRecordSetTypes.mcortAll) & " FROM manned_collectors mc WHERE " & mvEnv.Connection.WhereClause(vWhereFields))
      If vRecordSet.Fetch() = True Then
        InitFromRecordSet(pEnv, vRecordSet, MannedCollectorRecordSetTypes.mcortAll)
      Else
        InitClassFields()
        SetDefaults()
      End If
      vRecordSet.CloseRecordSet()
    End Sub
    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As MannedCollectorRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(MannedCollectorFields.mcofCollectorNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And MannedCollectorRecordSetTypes.mcortAll) = MannedCollectorRecordSetTypes.mcortAll Then
          .SetItem(MannedCollectorFields.mcofCollectionNumber, vFields)
          .SetItem(MannedCollectorFields.mcofContactNumber, vFields)
          .SetItem(MannedCollectorFields.mcofTotalTime, vFields)
          .SetItem(MannedCollectorFields.mcofAttended, vFields)
          .SetItem(MannedCollectorFields.mcofReadyForConfirmation, vFields)
          .SetItem(MannedCollectorFields.mcofReadyForAcknowledgement, vFields)
          .SetItem(MannedCollectorFields.mcofConfirmationProducedOn, vFields)
          .SetItem(MannedCollectorFields.mcofAcknowledgementProducedOn, vFields)
          .SetItem(MannedCollectorFields.mcofNotes, vFields)
          .SetItem(MannedCollectorFields.mcofAmendedBy, vFields)
          .SetItem(MannedCollectorFields.mcofAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(MannedCollectorFields.mcofAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    'Public Sub Create(ByVal pCollectorNumber As Long, ByVal pContactNumber As Long, Optional ByVal pTotalTime As String = "", Optional ByVal pNotes As String = "", Optional ByVal pAttended As String = "")
    '  With mvClassFields
    '    .Item(mcofCollectorNumber).Value = pCollectorNumber
    '    .Item(mcofContactNumber).Value = pContactNumber
    '    If Len(pTotalTime) > 0 Then .Item(mcofTotalTime).Value = pTotalTime
    '    If Len(pNotes) > 0 Then .Item(mcofNotes).Value = pNotes
    '    If Len(pAttended) > 0 Then .Item(mcofAttended).Bool = pAttended
    '  End With
    'End Sub

    Public Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      Init(pEnv)
      With mvClassFields
        .Item(MannedCollectorFields.mcofCollectionNumber).Value = pParams("CollectionNumber").Value
        .Item(MannedCollectorFields.mcofContactNumber).Value = pParams("ContactNumber").Value
        .Item(MannedCollectorFields.mcofAttended).Value = pParams("Attended").Value
        .Item(MannedCollectorFields.mcofReadyForConfirmation).Value = pParams("ReadyForConfirmation").Value
        .Item(MannedCollectorFields.mcofReadyForAcknowledgement).Value = pParams("ReadyForAcknowledgement").Value
        If pParams.Exists("TotalTime") Then .Item(MannedCollectorFields.mcofTotalTime).Value = pParams("TotalTime").Value
        If pParams.Exists("ConfirmationProducedOn") Then .Item(MannedCollectorFields.mcofConfirmationProducedOn).Value = pParams("ConfirmationProducedOn").Value
        If pParams.Exists("AcknowledgementProducedOn") Then .Item(MannedCollectorFields.mcofAcknowledgementProducedOn).Value = pParams("AcknowledgementProducedOn").Value
        If pParams.Exists("Notes") Then .Item(MannedCollectorFields.mcofNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Update(ByRef pParams As CDBParameters)
      'Auto Generated code for WEB services

      With mvClassFields
        If pParams.Exists("TotalTime") Then .Item(MannedCollectorFields.mcofTotalTime).Value = pParams("TotalTime").Value
        If pParams.Exists("Attended") Then .Item(MannedCollectorFields.mcofAttended).Value = pParams("Attended").Value
        If pParams.Exists("ReadyForConfirmation") Then .Item(MannedCollectorFields.mcofReadyForConfirmation).Value = pParams("ReadyForConfirmation").Value
        If pParams.Exists("ReadyForAcknowledgement") Then .Item(MannedCollectorFields.mcofReadyForAcknowledgement).Value = pParams("ReadyForAcknowledgement").Value
        If pParams.Exists("ConfirmationProducedOn") Then .Item(MannedCollectorFields.mcofConfirmationProducedOn).Value = pParams("ConfirmationProducedOn").Value
        If pParams.Exists("AcknowledgementProducedOn") Then .Item(MannedCollectorFields.mcofAcknowledgementProducedOn).Value = pParams("AcknowledgementProducedOn").Value
        If pParams.Exists("Notes") Then .Item(MannedCollectorFields.mcofNotes).Value = pParams("Notes").Value
      End With
    End Sub

    Public Sub Delete(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vDS As New VBDataSelection
      Dim vDT As CDBDataTable
      Dim vParams As CDBParameters

      vParams = New CDBParameters
      vParams.Add("ContactNumber", ContactNumber)
      vParams.Add("CollectionNumber", CollectionNumber)
      vDS.Init(mvEnv, DataSelection.DataSelectionTypes.dstContactCollectionPayments, vParams)
      vDT = vDS.DataTable
      If vDT.Rows.Count() > 0 Then RaiseError(DataAccessErrors.daeCannotDeleteCollectorHasPayments)
      vWhereFields.Add("collector_number", CDBField.FieldTypes.cftLong, CollectorNumber)
      vUpdateFields.Add("collector_number", CDBField.FieldTypes.cftLong)
      mvEnv.Connection.StartTransaction()
      mvClassFields.Delete(mvEnv.Connection, mvEnv, pAmendedBy, pAudit)
      mvEnv.Connection.UpdateRecords("collection_boxes", vUpdateFields, vWhereFields, False)
      mvEnv.Connection.DeleteRecords("collector_shifts", vWhereFields, False)
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

    Public ReadOnly Property AcknowledgementProducedOn() As String
      Get
        AcknowledgementProducedOn = mvClassFields.Item(MannedCollectorFields.mcofAcknowledgementProducedOn).Value
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(MannedCollectorFields.mcofAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(MannedCollectorFields.mcofAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Attended() As Boolean
      Get
        Attended = mvClassFields.Item(MannedCollectorFields.mcofAttended).Bool
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        Return mvClassFields.Item(MannedCollectorFields.mcofCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectorNumber() As Integer
      Get
        Return mvClassFields.Item(MannedCollectorFields.mcofCollectorNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ConfirmationProducedOn() As String
      Get
        ConfirmationProducedOn = mvClassFields.Item(MannedCollectorFields.mcofConfirmationProducedOn).Value
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields.Item(MannedCollectorFields.mcofContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(MannedCollectorFields.mcofNotes).Value
      End Get
    End Property

    Public ReadOnly Property ReadyForAcknowledgement() As Boolean
      Get
        ReadyForAcknowledgement = mvClassFields.Item(MannedCollectorFields.mcofReadyForAcknowledgement).Bool
      End Get
    End Property

    Public ReadOnly Property ReadyForConfirmation() As Boolean
      Get
        ReadyForConfirmation = mvClassFields.Item(MannedCollectorFields.mcofReadyForConfirmation).Bool
      End Get
    End Property

    Public ReadOnly Property TotalTime() As String
      Get
        TotalTime = mvClassFields.Item(MannedCollectorFields.mcofTotalTime).Value
      End Get
    End Property
  End Class
End Namespace

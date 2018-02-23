

Namespace Access
  Public Class Collector

    Public Enum CollectorRecordSetTypes 'These are bit values
      cllrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum CollectorFields
      cfAll = 0
      cfCollectorNumber
      cfCollectionNumber
      cfContactNumber
      cfTotalTime
      cfNotes
      cfAttended
      cfAmendedBy
      cfAmendedOn
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
          .DatabaseTableName = "collectors"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("collector_number", CDBField.FieldTypes.cftLong)
          .Add("collection_number", CDBField.FieldTypes.cftLong)
          .Add("contact_number", CDBField.FieldTypes.cftLong)
          .Add("total_time", CDBField.FieldTypes.cftNumeric)
          .Add("notes")
          .Add("attended")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftDate)
        End With

        mvClassFields.Item(CollectorFields.cfCollectionNumber).SetPrimaryKeyOnly()
        mvClassFields.Item(CollectorFields.cfContactNumber).SetPrimaryKeyOnly()
      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As CollectorFields)
      'Add code here to ensure all values are valid before saving
      If mvClassFields.Item(CollectorFields.cfCollectorNumber).IntegerValue = 0 Then mvClassFields.Item(CollectorFields.cfCollectorNumber).IntegerValue = mvEnv.GetControlNumber("CE")
      mvClassFields.Item(CollectorFields.cfAmendedOn).Value = TodaysDate()
      mvClassFields.Item(CollectorFields.cfAmendedBy).Value = mvEnv.User.Logname
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As CollectorRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = CollectorRecordSetTypes.cllrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "c")
      End If
      Return vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pCollectionNumber As Integer = 0, Optional ByRef pContactNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pCollectionNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(CollectorRecordSetTypes.cllrtAll) & " FROM collectors c WHERE collection_number = " & pCollectionNumber & " AND contact_number = " & pContactNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, CollectorRecordSetTypes.cllrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As CollectorRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(CollectorFields.cfCollectorNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And CollectorRecordSetTypes.cllrtAll) = CollectorRecordSetTypes.cllrtAll Then
          .SetItem(CollectorFields.cfCollectionNumber, vFields)
          .SetItem(CollectorFields.cfContactNumber, vFields)
          .SetItem(CollectorFields.cfTotalTime, vFields)
          .SetItem(CollectorFields.cfNotes, vFields)
          .SetItem(CollectorFields.cfAttended, vFields)
          .SetItem(CollectorFields.cfAmendedBy, vFields)
          .SetItem(CollectorFields.cfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(CollectorFields.cfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    Public Sub Create(ByVal pCollectionNumber As Integer, ByVal pContactNumber As Integer, Optional ByVal pTotalTime As String = "", Optional ByVal pNotes As String = "", Optional ByVal pAttended As String = "")
      With mvClassFields
        .Item(CollectorFields.cfCollectionNumber).Value = CStr(pCollectionNumber)
        .Item(CollectorFields.cfContactNumber).Value = CStr(pContactNumber)
        If Len(pTotalTime) > 0 Then .Item(CollectorFields.cfTotalTime).Value = pTotalTime
        If Len(pNotes) > 0 Then .Item(CollectorFields.cfNotes).Value = pNotes
        If Len(pAttended) > 0 Then .Item(CollectorFields.cfAttended).Bool = CBool(pAttended)
      End With
    End Sub

    Public Sub Update(ByVal pTotalTime As String, ByVal pNotes As String, ByVal pAttended As String)
      With mvClassFields
        .Item(CollectorFields.cfTotalTime).DoubleValue = CDbl(pTotalTime)
        .Item(CollectorFields.cfNotes).Value = pNotes
        .Item(CollectorFields.cfAttended).Value = pAttended
      End With
    End Sub

    Public Sub Delete()
      mvEnv.Connection.DeleteRecords("collectors", mvClassFields.WhereFields)
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
        AmendedBy = mvClassFields.Item(CollectorFields.cfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(CollectorFields.cfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property Attended() As String
      Get
        Attended = mvClassFields.Item(CollectorFields.cfAttended).Value
      End Get
    End Property

    Public ReadOnly Property CollectionNumber() As Integer
      Get
        CollectionNumber = mvClassFields.Item(CollectorFields.cfCollectionNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property CollectorNumber() As Integer
      Get
        CollectorNumber = mvClassFields.Item(CollectorFields.cfCollectorNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property ContactNumber() As Integer
      Get
        ContactNumber = mvClassFields.Item(CollectorFields.cfContactNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(CollectorFields.cfNotes).Value
      End Get
    End Property

    Public ReadOnly Property TotalTime() As Double
      Get
        TotalTime = mvClassFields.Item(CollectorFields.cfTotalTime).DoubleValue
      End Get
    End Property
  End Class
End Namespace
